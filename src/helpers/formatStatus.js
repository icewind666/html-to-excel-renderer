var formatStatus = function () {};

formatStatus.register = function (Handlebars) {
    Handlebars.registerHelper('formatStatus', function(status) {
        return status ? 'Активен' : 'Не активен';
    });
};

module.exports = formatStatus;