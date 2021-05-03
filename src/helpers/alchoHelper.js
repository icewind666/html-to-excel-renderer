var alchoHelper = function () {};

alchoHelper.register = function (Handlebars) {
    Handlebars.registerHelper('alchoHelper', function(step, value) {
        if (step === 0 && value === '0.0') {
            return '-';
        }
        if (step === 1 && value !== '0.0') {
            return 'Обнаружен';
        }
        return value;
    });
};

module.exports = alchoHelper;