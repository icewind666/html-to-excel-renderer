var dashOrData = function () {};

dashOrData.register = function (Handlebars) {
    Handlebars.registerHelper('dashOrData', function(data) {
        return data === null ? '-' : data;
    });
};

module.exports = dashOrData;