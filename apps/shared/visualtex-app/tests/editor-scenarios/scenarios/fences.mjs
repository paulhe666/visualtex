// nested fences and their delimiters

export const names = ["fences"];

export async function run({ clearField, evaluate, key }) {
  await clearField();
  const readFenceState = async (name) => ({
    name,
    ...(await evaluate(`(() => {
          const field = document.querySelector("math-field");
          const model = field?._mathfield?.model;
          const position = field?.position ?? -1;
          const around = model
            ? Array.from(
                { length: model.lastOffset + 1 },
                (_, offset) => {
                  const atom = model.at(offset);
                  return {
                    offset,
                    type: atom?.type ?? "",
                    branch: Array.isArray(atom?.parentBranch)
                      ? JSON.stringify(atom.parentBranch)
                      : atom?.parentBranch ?? null,
                    parentType: atom?.parent?.type ?? null,
                    latex: field.getElementInfo(offset)?.latex ?? "",
                  };
                },
              ).filter((item) => Math.abs(item.offset - position) <= 3)
            : [];
          return {
            value: field?.value ?? "",
            position,
            lastOffset: field?.lastOffset ?? -1,
            selection: field?.selection ?? null,
            around,
          };
        })()`)),
  });
  const states = [];
  await key("(", "Digit9", 57);
  states.push(await readFenceState("outer-open"));
  await key("(", "Digit9", 57);
  states.push(await readFenceState("inner-open"));
  await key("x", "KeyX", 88);
  states.push(await readFenceState("inner-content"));
  await key(")", "Digit0", 48);
  states.push(await readFenceState("inner-close"));
  await key("y", "KeyY", 89);
  states.push(await readFenceState("between-fences"));
  await key(")", "Digit0", 48);
  states.push(await readFenceState("outer-close"));
  await key("z", "KeyZ", 90);
  states.push(await readFenceState("after-fences"));

  const expectedValues = [
    String.raw`\left(\right)`,
    String.raw`\left(\left(\right)\right)`,
    String.raw`\left(\left(x\right)\right)`,
    String.raw`\left(\left(x\right)\right)`,
    String.raw`\left(\left(x\right)y\right)`,
    String.raw`\left(\left(x\right)y\right)`,
    String.raw`\left(\left(x\right)y\right)z`,
  ];
  states.forEach((state, index) => {
    if (state.value !== expectedValues[index]) {
      throw new Error(
        `Nested fence step ${state.name} changed structure: ${JSON.stringify({ state, expected: expectedValues[index] })}`,
      );
    }
  });
  if (
    !(states[3].position < states[3].lastOffset) ||
    !(states[5].position === states[5].lastOffset)
  ) {
    throw new Error(
      `Nested fence closing did not exit exactly one level at a time: ${JSON.stringify(states)}`,
    );
  }

  await clearField();
  await key("(", "Digit9", 57);
  await key("(", "Digit9", 57);
  await key("q", "KeyQ", 81);
  await key("^", "Digit6", 54);
  await key("a", "KeyA", 65);
  const scriptExitInsideFence = await readFenceState("script-exit-inside-inner-fence");
  await key(")", "Digit0", 48);
  const afterInnerClose = await readFenceState("after-inner-close");
  await key("b", "KeyB", 66);
  const afterInnerFence = await readFenceState("after-inner-fence");
  await key(")", "Digit0", 48);
  await key("c", "KeyC", 67);
  const afterOuterFence = await readFenceState("after-outer-fence");
  if (
    scriptExitInsideFence.value !== String.raw`\left(\left(q^{a}\right)\right)` ||
    afterInnerFence.value !== String.raw`\left(\left(q^{a}\right)b\right)` ||
    afterOuterFence.value !== String.raw`\left(\left(q^{a}\right)b\right)c`
  ) {
    throw new Error(
      `Script auto-exit and nested fence ownership conflicted: ${JSON.stringify({ scriptExitInsideFence, afterInnerClose, afterInnerFence, afterOuterFence })}`,
    );
  }
  states.push(
    scriptExitInsideFence,
    afterInnerClose,
    afterInnerFence,
    afterOuterFence,
  );
  console.log(JSON.stringify(states, null, 2));
  console.log("Targeted nested fence regression passed");
  return;
}
