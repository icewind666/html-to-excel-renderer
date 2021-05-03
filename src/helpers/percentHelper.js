var percentHelper = function () {};

percentHelper.register = function (Handlebars) {
    Handlebars.registerHelper('percentHelper', function(num) {
        return `${(num * 100).toFixed(2)}%`;
    });
};

module.exports = percentHelper;