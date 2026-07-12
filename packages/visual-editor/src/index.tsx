import { useEffect, useRef } from "react";
import type { NodeKind, SupportLevel, VisualNode } from "@visualtex/protocol";
import { DOMParser as ProseMirrorDOMParser, Schema, type Node as ProseMirrorNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";
import { EditorView, type NodeView } from "prosemirror-view";

export interface VisualEditorProps {
  nodes: VisualNode[];
  revision: number;
  onNodeEdit: (nodeId: string, content: string) => void;
  onNodeSelect?: (node: VisualNode | null) => void;
  onUndo?: () => void;
  onRedo?: () => void;
}

const schema = new Schema({
  nodes: {
    doc: { content: "visual_block*" },
    visual_block: {
      group: "block",
      content: "text*",
      defining: true,
      attrs: {
        nodeId: { default: "" },
        kind: { default: "paragraph" },
        support: { default: "native" },
        editable: { default: true },
      },
      parseDOM: [{ tag: "div[data-node-id]" }],
      toDOM(node) {
        return [
          "div",
          {
            "data-node-id": node.attrs.nodeId as string,
            "data-kind": node.attrs.kind as string,
            "data-support": node.attrs.support as string,
            class: `vt-visual-block vt-kind-${node.attrs.kind as string}`,
          },
          0,
        ];
      },
    },
    text: { group: "inline" },
  },
});

class VisualBlockView implements NodeView {
  dom: HTMLElement;
  contentDOM: HTMLElement;

  constructor(node: ProseMirrorNode) {
    this.dom = document.createElement("article");
    this.dom.className = `vt-visual-block vt-kind-${String(node.attrs.kind)}`;
    this.dom.dataset.nodeId = String(node.attrs.nodeId);
    this.dom.dataset.kind = String(node.attrs.kind);
    this.dom.dataset.support = String(node.attrs.support);

    const label = document.createElement("div");
    label.className = "vt-visual-label";
    label.textContent = labelForKind(node.attrs.kind as NodeKind, node.attrs.support as SupportLevel);
    this.dom.appendChild(label);

    this.contentDOM = document.createElement("div");
    this.contentDOM.className = "vt-visual-content";
    if (!node.attrs.editable) {
      this.contentDOM.contentEditable = "false";
      this.dom.classList.add("is-readonly");
    }
    this.dom.appendChild(this.contentDOM);
  }

  update(node: ProseMirrorNode): boolean {
    if (node.type !== schema.nodes.visual_block) return false;
    this.dom.dataset.kind = String(node.attrs.kind);
    this.dom.dataset.support = String(node.attrs.support);
    return true;
  }
}

const labelForKind = (kind: NodeKind, support: SupportLevel): string => {
  const labels: Partial<Record<NodeKind, string>> = {
    title: "标题",
    author: "作者",
    abstract: "摘要",
    section: "章节",
    subsection: "小节",
    paragraph: "正文",
    inline_math: "行内公式",
    display_math: "行间公式",
    figure: "图片",
    table: "表格",
    list: "列表",
    theorem: "定理环境",
    citation: "引用",
    reference: "交叉引用",
    footnote: "脚注",
    bibliography: "参考文献",
    raw_latex: "原始 LaTeX",
  };
  const suffix = support === "unstable" ? " · 语法未完成" : support === "opaque" ? " · 源码模式" : "";
  return `${labels[kind] ?? kind}${suffix}`;
};

const visibleNodes = (nodes: VisualNode[]): VisualNode[] =>
  nodes.filter((node) => node.kind !== "document" && node.kind !== "preamble" && node.text !== null);

const nodeIsEditable = (node: VisualNode): boolean =>
  node.support !== "opaque" &&
  node.support !== "unstable" &&
  !["raw_latex", "document", "preamble"].includes(node.kind);

function createDocument(nodes: VisualNode[]): ProseMirrorNode {
  return schema.node(
    "doc",
    null,
    visibleNodes(nodes).map((node) =>
      schema.node(
        "visual_block",
        {
          nodeId: node.id,
          kind: node.kind,
          support: node.support,
          editable: nodeIsEditable(node),
        },
        node.text ? schema.text(node.text) : undefined,
      ),
    ),
  );
}

function contentById(documentNode: ProseMirrorNode): Map<string, string> {
  const result = new Map<string, string>();
  documentNode.forEach((node) => {
    result.set(String(node.attrs.nodeId), node.textContent);
  });
  return result;
}

function selectedNodeId(state: EditorState): string | null {
  const resolved = state.selection.$from;
  if (resolved.depth < 1) return null;
  const node = resolved.node(1);
  return node.type === schema.nodes.visual_block ? String(node.attrs.nodeId) : null;
}

export function VisualEditor({
  nodes,
  revision,
  onNodeEdit,
  onNodeSelect,
  onUndo,
  onRedo,
}: VisualEditorProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const viewRef = useRef<EditorView | null>(null);
  const suppressRef = useRef(false);
  const propsRef = useRef({ nodes, onNodeEdit, onNodeSelect, onUndo, onRedo });
  propsRef.current = { nodes, onNodeEdit, onNodeSelect, onUndo, onRedo };

  useEffect(() => {
    if (!hostRef.current) return;
    const state = EditorState.create({ doc: createDocument(propsRef.current.nodes) });
    let view: EditorView;
    view = new EditorView(hostRef.current, {
      state,
      nodeViews: {
        visual_block: (node) => new VisualBlockView(node),
      },
      dispatchTransaction(transaction: Transaction) {
        const before = contentById(view.state.doc);
        const nextState = view.state.apply(transaction);
        view.updateState(nextState);
        if (!suppressRef.current && transaction.docChanged) {
          const after = contentById(nextState.doc);
          for (const [nodeId, content] of after) {
            if (before.get(nodeId) !== content) {
              propsRef.current.onNodeEdit(nodeId, content);
            }
          }
        }
        const selected = selectedNodeId(nextState);
        propsRef.current.onNodeSelect?.(
          propsRef.current.nodes.find((node) => node.id === selected) ?? null,
        );
      },
      handleKeyDown(_view, event) {
        if (!(event.metaKey || event.ctrlKey)) return false;
        const key = event.key.toLowerCase();
        if (key === "z" && event.shiftKey) {
          propsRef.current.onRedo?.();
          return Boolean(propsRef.current.onRedo);
        }
        if (key === "z") {
          propsRef.current.onUndo?.();
          return Boolean(propsRef.current.onUndo);
        }
        if (key === "y") {
          propsRef.current.onRedo?.();
          return Boolean(propsRef.current.onRedo);
        }
        return false;
      },
    });
    viewRef.current = view;
    return () => {
      view.destroy();
      viewRef.current = null;
    };
  }, []);

  useEffect(() => {
    const view = viewRef.current;
    if (!view) return;
    const current = contentById(view.state.doc);
    const expected = visibleNodes(nodes);
    const unchanged =
      current.size === expected.length &&
      expected.every((node) => current.get(node.id) === (node.text ?? ""));
    if (unchanged) return;

    suppressRef.current = true;
    const nextState = EditorState.create({ doc: createDocument(nodes) });
    view.updateState(nextState);
    suppressRef.current = false;
  }, [nodes, revision]);

  return <div ref={hostRef} className="vt-visual-editor" aria-label="Structured visual editor" />;
}

export { schema as visualEditorSchema, ProseMirrorDOMParser };
