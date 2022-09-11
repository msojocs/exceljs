const _ = require('../../../utils/under-dash');
const defaultNumFormats = require('../../defaultnumformats');

const BaseXform = require('../base-xform');

function hashDefaultFormats() {
  const hash = {};
  _.each(defaultNumFormats, (dnf, id) => {
    if (dnf.f) {
      hash[`f*${dnf.f}`] = parseInt(id, 10);
      // hash[`${dnf.f}`] = parseInt(id, 10);
    }
    // at some point, add the other cultures here...
    else {
      Object.keys(dnf).forEach(key => {
        if (dnf.hasOwnProperty(key)) {
          hash[`${key}*${dnf[key]}`] = parseInt(id, 10);
        }
      });
    }
  });
  return hash;
}
const defaultFmtHash = hashDefaultFormats();

// NumFmt encapsulates translation between number format and xlsx
class NumFmtXform extends BaseXform {
  constructor(id, formatCode) {
    super();

    this.id = id;
    this.formatCode = formatCode;
  }

  get tag() {
    return 'numFmt';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('numFmt', {numFmtId: model.id, formatCode: model.formatCode});
  }

  parseOpen(node) {
    switch (node.name) {
      case 'numFmt':
        this.model = {
          id: parseInt(node.attributes.numFmtId, 10),
          formatCode: node.attributes.formatCode.replace(/[\\](.)/g, '$1'),
        };
        return true;
      default:
        return false;
    }
  }

  parseText() { }

  parseClose() {
    return false;
  }
}

NumFmtXform.getDefaultFmtId = function getDefaultFmtId(formatCode) {
  return defaultFmtHash[formatCode];
};

NumFmtXform.getDefaultFmtCode = function getDefaultFmtCode(numFmtId) {
  const locale = Intl.DateTimeFormat().resolvedOptions().locale.toLowerCase();
  // return defaultNumFormats[numFmtId] && (defaultNumFormats[numFmtId].f || defaultNumFormats[numFmtId][locale]);
  if(defaultNumFormats[numFmtId]){
    if(defaultNumFormats[numFmtId].f){
      return `f*${defaultNumFormats[numFmtId].f}`;
    }
    if(defaultNumFormats[numFmtId][locale]){
      return `${locale}*${defaultNumFormats[numFmtId][locale]}`;
    }
  }
  return null;
};

module.exports = NumFmtXform;
