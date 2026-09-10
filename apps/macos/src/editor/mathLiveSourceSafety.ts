export type MathLiveSourceSafetyIssueKind =
  | "source-length"
  | "group-depth"
  | "environment-depth";

export interface MathLiveSourceSafetyIssue {
  kind: MathLiveSourceSafetyIssueKind;
  value: number;
  limit: number;
}

// These limits protect the recursive MathLive 0.109 renderer used by WKWebView.
// In addition to deep nesting, sufficiently long flat atom sequences eventually
// exhaust the JS call stack and become very expensive well before that point.
// 64 KiB is far above ordinary single-formula source sizes while keeping a bad
// persisted row from blocking startup for seconds or overflowing the renderer.
export const MAX_MATHLIVE_SOURCE_LENGTH = 64 * 1024;
export const MAX_MATHLIVE_GROUP_DEPTH = 256;
export const MAX_MATHLIVE_ENVIRONMENT_DEPTH = 160;

function isEscaped(source: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

export function inspectMathLiveSourceSafety(
  source: string,
): MathLiveSourceSafetyIssue | null {
  if (source.length > MAX_MATHLIVE_SOURCE_LENGTH) {
    return {
      kind: "source-length",
      value: source.length,
      limit: MAX_MATHLIVE_SOURCE_LENGTH,
    };
  }

  let groupDepth = 0;
  let maximumGroupDepth = 0;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (character === "{" && !isEscaped(source, index)) {
      groupDepth += 1;
      maximumGroupDepth = Math.max(maximumGroupDepth, groupDepth);
      if (maximumGroupDepth > MAX_MATHLIVE_GROUP_DEPTH) {
        return {
          kind: "group-depth",
          value: maximumGroupDepth,
          limit: MAX_MATHLIVE_GROUP_DEPTH,
        };
      }
    } else if (character === "}" && !isEscaped(source, index)) {
      groupDepth = Math.max(0, groupDepth - 1);
    }
  }

  let environmentDepth = 0;
  let maximumEnvironmentDepth = 0;
  const environmentToken = /\\(begin|end)\s*\{[^{}]{1,64}\}/g;
  for (const match of source.matchAll(environmentToken)) {
    const index = match.index ?? 0;
    if (isEscaped(source, index)) continue;
    if (match[1] === "begin") {
      environmentDepth += 1;
      maximumEnvironmentDepth = Math.max(
        maximumEnvironmentDepth,
        environmentDepth,
      );
      if (maximumEnvironmentDepth > MAX_MATHLIVE_ENVIRONMENT_DEPTH) {
        return {
          kind: "environment-depth",
          value: maximumEnvironmentDepth,
          limit: MAX_MATHLIVE_ENVIRONMENT_DEPTH,
        };
      }
    } else {
      environmentDepth = Math.max(0, environmentDepth - 1);
    }
  }

  return null;
}

export function mathLiveSourceSafetyMessage(
  issue: MathLiveSourceSafetyIssue,
  language: "cn" | "en" = "cn",
) {
  if (language === "en") {
    if (issue.kind === "source-length") {
      return `Formula source is too large to render safely (${issue.value} characters; limit ${issue.limit}).`;
    }
    if (issue.kind === "group-depth") {
      return `Formula group nesting is too deep to render safely (depth ${issue.value}; limit ${issue.limit}).`;
    }
    return `Formula environment nesting is too deep to render safely (depth ${issue.value}; limit ${issue.limit}).`;
  }

  if (issue.kind === "source-length") {
    return `公式源码过大，已阻止渲染以避免界面崩溃（${issue.value} 字符，上限 ${issue.limit}）。`;
  }
  if (issue.kind === "group-depth") {
    return `公式分组嵌套过深，已阻止渲染以避免界面崩溃（深度 ${issue.value}，上限 ${issue.limit}）。`;
  }
  return `公式环境嵌套过深，已阻止渲染以避免界面崩溃（深度 ${issue.value}，上限 ${issue.limit}）。`;
}
