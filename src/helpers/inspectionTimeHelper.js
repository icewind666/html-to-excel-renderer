var inspectionTimeHelper = function () {};

inspectionTimeHelper.register = function (Handlebars) {
    Handlebars.registerHelper('inspectionTimeHelper', function(min = '00', sec = '00') {
        return `${min}:${sec}`;
    });
};

module.exports = inspectionTimeHelper;