var formatComplains = function () {};

formatComplains.register = function (Handlebars) {
    Handlebars.registerHelper('formatComplains', function(complains) {
        if (complains === null) {
            return '-';
        }

        return complains ? 'Есть' : 'Нет';
    });
};

module.exports = formatComplains;