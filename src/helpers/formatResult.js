var formatResult = function () {};

formatResult.register = function (Handlebars) {
    Handlebars.registerHelper('formatResult', function(result) {
        return result ? 'Допуск' : 'Не допуск';
    });
};

module.exports = formatResult;