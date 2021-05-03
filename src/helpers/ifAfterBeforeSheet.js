var isAfterBeforeSheet = function () {};

isAfterBeforeSheet.register = function (Handlebars) {
    Handlebars.registerHelper('isAfterBeforeSheet', function(sheetName) {
        return sheetName.includes('Предрейсовый')
            || sheetName.includes('Послерейсовый')
            || sheetName.includes('Линейный');
    });
};

module.exports = isAfterBeforeSheet;