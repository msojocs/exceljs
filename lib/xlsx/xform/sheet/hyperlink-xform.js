const BaseXform = require('../base-xform');

class HyperlinkXform extends BaseXform {
  get tag() {
    return 'hyperlink';
  }

  render(xmlStream, model) {

    // console.log('model:', model);
    // throw new Error(123)
    if (this.isInternalLink(model)) {
      xmlStream.leafNode('hyperlink', {
        ref: model.address,
        tooltip: model.tooltip,
        location: model.location,
      });
    } else {
      xmlStream.leafNode('hyperlink', {
        ref: model.address,
        'r:id': model.rId,
        tooltip: model.tooltip,
      });
    }
  }

  parseOpen(node) {
    if (node.name === 'hyperlink') {
      this.model = {
        address: node.attributes.ref,
        tooltip: node.attributes.tooltip,
      };

      // This is an internal link
      if (node.attributes.location) {
        this.model.location = node.attributes.location;
      }
      if(node.attributes['r:id']){
        this.model.rId = node.attributes['r:id'];
      }
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }

  // @Deprecated https://google.com/AB:A3 is ok too
  isInternalLink(model) {
    // @example: Sheet2!D3, return true
    return model.location && /^[^!]+![a-zA-Z]+[\d]+$/.test(model.location);
  }
}

module.exports = HyperlinkXform;
