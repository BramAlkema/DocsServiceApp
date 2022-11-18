// --- PowerpointApp ---
(function(r) {
  var PowerpointApp;
  PowerpointApp = (function() {
    var disassemblePowerpoint, getXmlObj, parsPPTX, putError, putInternalError;

    class PowerpointApp {
      constructor(blob_) {
        this.name = "PowerpointApp";
        if (!blob_ || blob_.getContentType() !== MimeType.MICROSOFT_POWERPOINT) {
          throw new Error("Please set the blob of data of PPTX format.");
        }
        this.obj = {
          powerpoint: blob_
        };
        this.contentTypes = "[Content_Types].xml";
        this.document = "ppt/presentation.xml";
        this.mainObj = {};
        parsPPTX.call(this);
      }

      // --- begin methods
      getTableColumnWidth() {
        var body, n1, obj, root, xmlObj;
        if (this.mainObj.fileObj.hasOwnProperty(this.document)) {
          xmlObj = getXmlObj.call(this, this.document);
          root = xmlObj.getRootElement();
          n1 = root.getNamespace("w");
          body = root.getChild("body", n1).getChildren("tbl", n1);
          obj = body.map((e, i) => {
            var tblGrid, tblPr, tblW, temp, w;
            temp = {
              tableIndex: i,
              unit: "pt"
            };
            tblPr = e.getChild("tblPr", n1);
            if (tblPr) {
              tblW = tblPr.getChild("tblW", n1);
              if (tblW) {
                w = tblW.getAttribute("w", n1);
                if (w) {
                  temp.tableWidth = Number(w.getValue()) / 20;
                }
              }
            }
            tblGrid = e.getChild("tblGrid", n1);
            if (tblGrid) {
              temp.tebleColumnWidth = tblGrid.getChildren("gridCol", n1).map((f) => {
                return Number(f.getAttribute("w", n1).getValue()) / 20;
              });
            }
            return temp;
          });
          return obj;
        }
      }

    };

    PowerpointApp.name = "PowerpointApp";

    // --- end methods
    parsPPTX = function() {
      disassemblePowerpoint.call(this);
    };

    disassemblePowerpoint = function() {
      var blobs;
      blobs = Utilities.unzip(this.obj.powerpoint.setContentType(MimeType.ZIP));
      this.mainObj.fileObj = blobs.reduce((o, b) => {
        return Object.assign(o, {
          [b.getName()]: b
        });
      }, {});
    };

    getXmlObj = function(k_) {
      return XmlService.parse(this.mainObj.fileObj[k_].getDataAsString());
    };

    putError = function(m) {
      throw new Error(`${m}`);
    };

    putInternalError = function(m) {
      throw new Error(`Internal error: ${m}`);
    };

    return PowerpointApp;

  }).call(this);
  return r.PowerpointApp = PowerpointApp;
})(this);
