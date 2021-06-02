var pressureHelper = function () {};

pressureHelper.register = function (Handlebars) {
    Handlebars.registerHelper('pressureHelper', function(pressure, upper) {
        if (pressure && upper && upper !== 0) {
            return pressure;
        }
        return '-';
    });
};

module.exports = pressureHelper;