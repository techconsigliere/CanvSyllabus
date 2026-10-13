// Minimal stand-in for the "option" package (some/none), covering what lop and mammoth call.
function callOrReturn(v) { return typeof v === "function" ? v() : v; }
var none = {
  isNone: function () { return true; }, isSome: function () { return false; },
  map: function () { return none; }, flatMap: function () { return none; },
  filter: function () { return none; }, toArray: function () { return []; },
  value: function () { throw new Error("Called value on none"); },
  valueOrElse: function (other) { return callOrReturn(other); },
  orElse: function (other) { return callOrReturn(other); }
};
function some(value) {
  return {
    isNone: function () { return false; }, isSome: function () { return true; },
    map: function (f) { return some(f(value)); }, flatMap: function (f) { return f(value); },
    filter: function (p) { return p(value) ? this : none; }, toArray: function () { return [value]; },
    value: function () { return value; }, valueOrElse: function () { return value; },
    orElse: function () { return this; }
  };
}
exports.none = none; exports.some = some;
exports.isOption = function (v) { return v === none || (v && typeof v.isSome === "function"); };
