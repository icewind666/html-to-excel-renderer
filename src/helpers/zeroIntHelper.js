var zeroIntHelper = function () {};

zeroIntHelper.register = function (Handlebars) {
    Handlebars.registerHelper('zeroIntHelper', function(value) {
        if (!value) {
            return '00';
        }
        if (Number(value) < 10) {
            return `0${value}`;
        }
        return value;

    });
};

module.exports = zeroIntHelper;