var formatDateTime = function () {};

formatDateTime.register = function (Handlebars) {
    Handlebars.registerHelper('formatDateTime', function(dateTime) {
        const date = new Date(dateTime);
        return `${date.getDate()}-${date.getMonth() + 1 }-${date.getFullYear()} \n${date.getHours()}:${date.getMinutes()}`;
    });
};

module.exports = formatDateTime;