var summarize = function () {};

summarize.register = function (Handlebars) {
    Handlebars.registerHelper('summarize', function(obj) {
        if (obj.sheetName.includes('Предрейсовый')) {
            return beforeSumRows(obj);
        }
        if (obj.sheetName.includes('Послерейсовый')) {
            return afterSumRows(obj);
        }
        return '';
    });
};

function beforeSumRows (obj) {
    const allCount = obj.inspections.length;
    let goodBefore = 0;
    let badBefore = 0;
    let goodLine = 0;
    let badLine = 0;
    for (const insp of obj.inspections) {
        if (insp.type === 'Предрейсовый' || insp.type === 'Предсменный') {
            if (insp.allow === 'Допущен') {
                goodBefore++;
            } else {
                badBefore++;
            }
        } else if (insp.allow === 'Допущен') {
            goodLine++;
        } else {
            badLine++;
        }
    }
    return `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого осмотрено: </td><td></td><td></td><td></td><td>${allCount}</td></tr>`
        + `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого допущено к исполнению трудовых обязанностей: </td><td></td><td></td><td></td><td>${goodBefore}</td></tr>`
        + `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого не допущено к исполнению трудовых обязанностей: </td><td></td><td></td><td></td><td>${badBefore}</td></tr>`;
    /* '<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого прошло линейный осмотр: </td><td></td><td></td><td></td><td>'+goodLine+'</td></tr>' +
          '<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого отстраненных от трудовых обязанностей: </td><td></td><td></td><td></td><td>'+badLine+'</td></tr>' */
}

function afterSumRows (obj) {
    const allCount = obj.inspections.length;
    let goodAfter = 0;
    let badAfter = 0;
    let goodAftershift = 0;
    let badAftershift = 0;
    for (const insp of obj.inspections) {
        if (['Послерейсовый'].includes(insp.type)) {
            if (insp.allow === 'Допущен') {
                goodAfter++;
            } else {
                badAfter++;
            }
        } else if (insp.allow === 'Допущен') {
            goodAftershift++;
        } else {
            badAftershift++;
        }
    }
    return `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого осмотрено: </td><td></td><td></td><td></td><td>${allCount}</td></tr>`
        + `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого прошло послерейсовый: </td><td></td><td></td><td></td><td>${goodAfter}</td></tr>`
        + `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого прошло послесменный: </td><td></td><td></td><td></td><td>${goodAftershift}</td></tr>`
        + `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого не прошло послерейсовый: </td><td></td><td></td><td></td><td>${badAfter}</td></tr>`
        + `<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого не прошло послесменный: </td><td></td><td></td><td></td><td>${badAftershift}</td></tr>`;
}

module.exports = summarize;