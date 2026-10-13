// mammoth only uses xmlbuilder to write a style map back into a .docx, which CanvSyllabus never does.
exports.create = function () { throw new Error("xmlbuilder is not included in this browser build"); };
