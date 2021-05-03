var lineSumRows = function () {};

lineSumRows.register = function (Handlebars) {
    Handlebars.registerHelper('lineSumRows', function(obj) {
        const allCount = obj.inspections.length;
        let goodLine = 0;
        let badLine = 0;
        for (const insp of obj.inspections) {
            if (insp.allow === 'Допущен') {
                goodLine++;
            } else {
                badLine++;
            }
        }
        return `<tr style="height: 300px; border-style: solid"><td colspan="4" style="text-align: left">Итого осмотрено: </td><td></td><td></td><td></td><td>${allCount}</td></tr>`
            + `<tr style="height: 300px; border-style: solid"><td colspan="4" style="text-align: left">Итого прошло линейный контроль: </td><td></td><td></td><td></td><td>${goodLine}</td></tr>`
            + `<tr style="height: 300px; border-style: solid"><td colspan="4" style="text-align: left">Итого отстраненных от трудовых обязанностей: </td><td></td><td></td><td></td><td>${badLine}</td></tr>`;
    });
};

module.exports = lineSumRows;