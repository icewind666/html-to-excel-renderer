var ifnull = function () {};

ifnull.register = function (Handlebars) {
    Handlebars.registerHelper('ifnull', function(obj, ifNull) {
        return obj || ifNull;
    });
};

module.exports = ifnull;