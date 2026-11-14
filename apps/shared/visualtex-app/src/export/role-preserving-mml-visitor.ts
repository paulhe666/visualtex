import { SerializedMmlVisitor } from 'mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js';
import { TEXCLASS, type MmlNode } from 'mathjax-full/js/core/MmlTree/MmlNode.js';

// Keep the parser's actual named-operator role when the stock serializer would
// omit it based on spelling. An autoOP=false font scope is not an ORD proof.
export class RolePreservingMmlVisitor extends SerializedMmlVisitor {
  protected getDataAttributes(node: MmlNode): {[name: string]: string} {
    const data = super.getDataAttributes(node);
    if (node.isKind('mi') && node.getProperty('texClass') === TEXCLASS.OP)
      data['data-mjx-texclass'] = 'OP';
    return data;
  }
}
