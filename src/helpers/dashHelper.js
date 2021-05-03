var dashHelper = function () {};

dashHelper.register = function (Handlebars) {
    Handlebars.registerHelper('dashHelper', function(value = '-') {
        if (value === '0' || value === 0) {
            return '-';
        }
        return value;
    });
};

module.exports = dashHelper;