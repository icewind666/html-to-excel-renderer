var dashMgtHelper = function () {};

dashMgtHelper.register = function (Handlebars) {
    Handlebars.registerHelper('dashMgtHelper', function(value) {
        if (value === '0/0' || value === '0.0') {
            return '-';
        }
        return value;
    });
};

module.exports = dashMgtHelper;