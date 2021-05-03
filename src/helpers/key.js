var key = function () {};

key.register = function (Handlebars) {
    Handlebars.registerHelper('key', function(obj, key) {
        return obj[key];
    });
};

module.exports = key;