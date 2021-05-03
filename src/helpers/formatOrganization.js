var formatOrganization = function () {};

formatOrganization.register = function (Handlebars) {
    Handlebars.registerHelper('formatOrganization', function(org) {
        return org && org.name ? org.name : '';
    });
};

module.exports = formatOrganization;