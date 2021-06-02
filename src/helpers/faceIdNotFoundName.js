var faceIdNotFoundName = function () {};

faceIdNotFoundName.register = function (Handlebars) {
    Handlebars.registerHelper('faceIdNotFoundName', function(name, surname, patronymic) {
        if (name) {
            return `${surname} ${name} ${patronymic}`;
        }
        return 'нет соответствия';
    });
};

module.exports = faceIdNotFoundName;