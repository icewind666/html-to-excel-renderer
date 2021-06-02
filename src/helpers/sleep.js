var sleep = function () {};

sleep.register = function (Handlebars) {
    Handlebars.registerHelper('sleep', function(sleep) {
        if (sleep === null) return '-'
        return sleep ? 'более 8 часов' : 'менее 8 часов'
    });
};

module.exports = sleep;