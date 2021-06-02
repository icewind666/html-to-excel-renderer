var formatType = function () {};

formatType.register = function (Handlebars) {
    Handlebars.registerHelper('formatType', function(type) {
        let result;
        switch (type) {
            case 'BEFORE': {
                result = 'Предрейсовый';
                break;
            }
            case 'BEFORE_SHIFT': {
                result = 'Предсменный';
                break;
            }
            case 'LINE': {
                result = 'Линейный';
                break;
            }
            case 'AFTER': {
                result = 'Послерейсовый';
                break;
            }
            case 'AFTER_SHIFT': {
                result = 'Послесменный';
                break;
            }
            case 'ALCO': {
                result = 'Алкотестирование';
                break;
            }
            case 'PIRO': {
                result = 'Контроль температуры';
                break;
            }
            case 'PREVENTION': {
                result = 'Профилактический';
                break;
            }
            default: {
                result = 'Неизвестный';
            }
        }

        return result;
    });
};

module.exports = formatType;