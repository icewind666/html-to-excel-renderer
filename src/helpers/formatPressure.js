var formatPressure = function () {};

formatPressure.register = function (Handlebars) {
    Handlebars.registerHelper('formatPressure', function(meddata) {
        return `${dashOrData(meddata.systolicPressure)} / ${dashOrData(meddata.diastolicPressure)}`;
    });
};

module.exports = formatPressure;