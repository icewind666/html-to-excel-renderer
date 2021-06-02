var allowHelper = function () {};

allowHelper.register = function (Handlebars) {
    Handlebars.registerHelper('allowHelper', function(allow) {
        if (allow === 'Допущен') {
            return 'Прошел';
        }
        return 'Не прошел';
    });
};

module.exports = allowHelper;