var formatGender = function () {};

formatGender.register = function (Handlebars) {
    Handlebars.registerHelper('formatGender', function(gender) {
        return gender === 'MALE' ? 'М' : 'Ж';
    });
};

module.exports = formatGender;