var upper = function () {};

upper.register = function (Handlebars) {
    Handlebars.registerHelper('upper', function(str) {
            return typeof str === 'string' ? str.toUpperCase() : str;
    });
};

module.exports = upper;