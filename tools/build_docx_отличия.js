// Сборка docs/Отличия_от_v4.docx — версии для Word.
//
// Источник содержания — docs/Отличия_от_v4.md; текст здесь продублирован
// намеренно: вёрстка Word (таблицы с фиксированными колонками, врезка,
// колонтитул) из разметки markdown не выводится, а конвертеры ломают
// кириллицу в именах файлов и ширины колонок.
//
//   npm install docx          # один раз, вне репозитория
//   node tools/build_docx_отличия.js docs/Отличия_от_v4.docx
//
// Проверка вида: soffice --headless --convert-to pdf (нужен
// libreoffice-writer, не только calc) и pdftoppm.

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, ShadingType, BorderStyle,
  TableLayoutType, VerticalAlign, Footer, PageNumber, LevelFormat,
} = require('docx');

const RED = '9E1B31', INK = '1A1A1A', MUTED = '595959', LINE = 'D4D4D4', BAND = 'F5F5F5';
const FONT = 'Arial';

// A4 (11906 x 16838 DXA), поля 2 см (1134) -> ширина текста 9638
const W = 9638;

const sz = (pt) => pt * 2;

function t(text, o = {}) {
  return new TextRun({
    text, font: FONT, size: sz(o.size || 10),
    bold: !!o.bold, italics: !!o.italics, color: o.color || INK,
  });
}

function p(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.LEFT,
    spacing: { before: (o.before ?? 0) * 20, after: (o.after ?? 6) * 20, line: 260 },
    indent: o.indent ? { left: o.indent } : undefined,
    children: Array.isArray(text) ? text : [t(text, o)],
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 340, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: RED, space: 4 } },
    children: [new TextRun({ text, font: FONT, size: sz(14), bold: true, color: RED })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, font: FONT, size: sz(11), bold: true, color: INK })],
  });
}

// ---- таблица ------------------------------------------------------------
// widths — доли, в сумме 1. align — 'l'|'r' по колонкам.
function table(widths, align, head, rows, o = {}) {
  const cols = widths.map((w) => Math.round(W * w));
  cols[cols.length - 1] = W - cols.slice(0, -1).reduce((a, b) => a + b, 0);

  const cell = (text, i, opts) => new TableCell({
    width: { size: cols[i], type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    },
    children: String(text).split('\n').map((line, k) => new Paragraph({
      alignment: align[i] === 'r' ? AlignmentType.RIGHT : AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: 240 },
      children: [new TextRun({
        text: line, font: FONT, size: sz(opts.size || 9),
        bold: !!opts.bold, color: opts.color || INK,
      })],
    })),
  });

  const шапка = new TableRow({
    tableHeader: true,
    children: head.map((x, i) => cell(x, i, { bold: true, fill: RED, color: 'FFFFFF' })),
  });

  const тело = rows.map((r, n) => new TableRow({
    children: r.map((x, i) => cell(x, i, {
      bold: o.boldLast && i === r.length - 1,
      fill: n % 2 ? BAND : undefined,
    })),
  }));

  return new Table({
    columnWidths: cols,
    width: { size: W, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    rows: [шапка, ...тело],
  });
}

const bullet = (text, lvl = 0) => new Paragraph({
  numbering: { reference: 'dots', level: lvl },
  spacing: { before: 0, after: 60, line: 260 },
  children: [t(text)],
});

const num = (text) => new Paragraph({
  numbering: { reference: 'nums', level: 0 },
  spacing: { before: 0, after: 60, line: 260 },
  children: Array.isArray(text) ? text : [t(text)],
});

// ---- содержание ---------------------------------------------------------
const children = [];

// Титульный блок
children.push(new Paragraph({
  spacing: { after: 40 },
  children: [new TextRun({ text: 'ФИНАНСОВАЯ МОДЕЛЬ УК «СПАРТАК»', font: FONT, size: sz(9), bold: true, color: MUTED })],
}));
children.push(new Paragraph({
  spacing: { after: 60 },
  children: [new TextRun({ text: 'Чем текущая модель отличается от v4.0', font: FONT, size: sz(20), bold: true, color: RED })],
}));
children.push(new Paragraph({
  spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: RED, space: 6 } },
  children: [new TextRun({
    text: 'Сравнение файла «Финансовая модель франшизы Спартак v4.0» (15 листов, около 28 400 формул) с текущей моделью Spartak_V9_0.xlsx (11 листов, 22 510 формул). Состояние на 11 сентября 2026 года.',
    font: FONT, size: sz(9.5), italics: true, color: MUTED,
  })],
}));

// Резюме
children.push(new Table({
  columnWidths: [W],
  width: { size: W, type: WidthType.DXA },
  layout: TableLayoutType.FIXED,
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      margins: { top: 160, bottom: 160, left: 200, right: 200 },
      shading: { type: ShadingType.CLEAR, fill: BAND, color: 'auto' },
      borders: {
        top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        left: { style: BorderStyle.SINGLE, size: 18, color: RED },
      },
      children: [
        p([t('Коротко. ', { bold: true }),
           t('Экономика франшизы осталась той же — паушальный взнос, роялти, мерч, доли ФК, займ 50 млн ₽. Изменились три вещи:')], { after: 8 }),
        p([t('из чего считается сеть', { bold: true }), t(' — партнёр перестал быть равен школе;')], { indent: 200, after: 4 }),
        p([t('на чём стоят доходные вводные', { bold: true }), t(' — факт реестра вместо круглых чисел;')], { indent: 200, after: 4 }),
        p([t('как устроена книга', { bold: true }), t(' — одно место ввода, один расчёт, проверяемый скриптом.')], { indent: 200, after: 0 }),
      ],
    })],
  })],
}));
children.push(p('', { after: 0 }));

// 1
children.push(h1('1. Структура файла'));
children.push(table([0.20, 0.40, 0.40], ['l', 'l', 'l'],
  ['', 'v4.0', 'Сейчас'],
  [
    ['Вводные', 'Размазаны по листам «Допущения», «План_продаж», «Типы_продукта», «Микс_продаж», «Штатное_расписание»', 'Все ручные вводные на листе 01_Вводные, больше нигде'],
    ['Сценарии', 'Три полных помесячных листа, считаются всегда', 'Один расчёт и переключатель на дашборде; сравнение трёх сценариев считается отдельно'],
    ['Именованные диапазоны', 'Нет', '203 имени — формулы читаются словами, а не адресами'],
    ['Правка модели', 'Вручную в ячейках', 'Скриптом: правка видна в диффе и откатывается'],
    ['Проверки', 'Лист «Проверка_модели», 54 формулы', 'Лист 08_Проверки, 55 автотестов и мутационная самопроверка на 13 подсаженных дефектов'],
  ]));
children.push(p('В v4 одну вводную приходилось менять в трёх-четырёх местах, и расхождение между сценарными листами ничем не ловилось. Сейчас вводная одна, а расхождение витрины и расчётного движка ловят проверки.', { before: 8, italics: true, color: MUTED, size: 9 }));

// 2
children.push(h1('2. Что изменилось в допущениях'));

children.push(h2('Календарь и план продаж'));
children.push(table([0.28, 0.36, 0.36], ['l', 'l', 'l'],
  ['', 'v4.0', 'Сейчас'],
  [
    ['Старт модели', '01.08.2026, сезон август — июль', '01.09.2026, сезон сентябрь — август; считается от любой даты старта'],
    ['План продаж, базовый', '37,8 / 43 / 43 / 43 / 43 = 210 сделок', '66 / 72 / 72 / 60 / 50 = 320 сделок'],
    ['План продаж, позитивный', '63,4 / 72 / 72 / 60 / 50 = 317 сделок', '—'],
  ]));
children.push(p([t('Текущий базовый план — это бывший позитивный сценарий v4.', { bold: true }), t(' Сравнивать сегодняшнюю базу корректно с колонкой «Позитивный» старого файла.')], { before: 8 }));

children.push(h2('Стартовая сеть'));
children.push(table([0.28, 0.36, 0.36], ['l', 'l', 'l'],
  ['', 'v4.0', 'Сейчас'],
  [
    ['Действующие школы', '51', '104 (по реестру сети)'],
    ['Действующие юрлица', 'Отдельно не считались: школа = юрлицо', '51'],
    ['Дети действующей сети', '51 × 100 = 5 100 (допущение)', '5 200 (по реестру)'],
  ]));

children.push(h2('Доходные вводные действующей сети'));
children.push(table([0.28, 0.36, 0.36], ['l', 'l', 'l'],
  ['', 'v4.0', 'Сейчас'],
  [
    ['Средний абонемент', '7 800 ₽', '6 000 ₽ (подтверждено владельцем)'],
    ['Минимальный роялти', '17 500 ₽ на школу, одним котлом', 'Пять групп точек начисления: только фикс, только процент, максимум из двух и так далее. Средний фикс 17 246 ₽ на точку'],
  ]));
children.push(p('Расчёт по группам воспроизводит фактический реестр начислений за август 2026 с точностью 0,1 %. Котёл «школы, умноженные на ставку» давал на том же реестре заметную ошибку в обе стороны.', { before: 8, italics: true, color: MUTED, size: 9 }));

children.push(h2('Мерч и склад'));
children.push(table([0.34, 0.33, 0.33], ['l', 'l', 'l'],
  ['', 'v4.0', 'Сейчас'],
  [
    ['Себестоимость комплекта', '4 400 ₽', '2 100 ₽'],
    ['Наценка', '70 %', '85 % (расчётная цена 3 885 ₽)'],
    ['Коэффициент формы', '1,2 комплекта на ребёнка за сезон', 'Без изменений'],
    ['Целевой запас', '4 месяца; вводная ни на что не влияла', '3 месяца будущих продаж, закупка до цели четыре раза в год'],
    ['Доля себестоимости на погашение займа', '50 %', '60 %'],
    ['Стартовый склад и займ ФК', '50 млн ₽ и 50 млн ₽', 'Без изменений'],
  ]));

children.push(h2('Ставки, налоги, доли ФК'));
children.push(table([0.40, 0.30, 0.30], ['l', 'l', 'l'],
  ['', 'v4.0', 'Сейчас'],
  [
    ['Роялти: действующие и продажи до 31.12.2026', '6,5 %', 'Без изменений'],
    ['Роялти: новые партнёры с 01.01.2027', '8,5 %', 'Без изменений'],
    ['Общее повышение до 10 %', 'с 01.10.2030', 'с 01.01.2030 (по ТЗ)'],
    ['Каникулы по роялти новых партнёров', '2 месяца', '0; штатный параметр B297'],
    ['Доли ФК: паушальный / роялти / прибыль мерча', '35 % / 10 % / 35 %', 'Без изменений'],
    ['НДС и УСН', '5 %; 15 % с минимумом 1 %', 'Без изменений'],
    ['Страховые взносы', 'Предельная база 2 979 000 ₽: 30 % до и 15,1 % сверх', 'Плоские 30 %, льгот нет (решение владельца)'],
    ['ERP', '180 000 ₽ при 60 школах и 3 000 ₽ за школу сверх', '80 000 ₽ при 104 школах и 3 000 ₽ за новую'],
    ['ФОТ ядра УК', '3 058 402 ₽ в месяц', 'Без изменений'],
  ]));

children.push(h2('Что перестало быть сценарным'));
children.push(p('В v4 сценарий двигал конверсию открытия (85 / 95 / 97 %), закрытия школ (5 / 3 / 2 %) и рост детей действующей сети (5 / 7 / 10 %). Сейчас это штатные вводные с одним значением для всех сценариев: конверсия 100 %, закрытия 3 % в год, рост 10 % в год.'));
children.push(p('Сценарий двигает объём продаж, набор детей и цену — то, что действительно неизвестно, а не механику расчёта.'));

// 3
children.push(h1('3. Что изменилось в логике'));

children.push(h2('Партнёр не равен школе — главное структурное изменение'));
children.push(p('В v4 одна продажа давала одну школу. Сейчас один партнёр открывает 1,5 школы, и не разом: первая — через лаг «продажа → открытие» (2 месяца), следующие — ещё через 3 месяца. Отсюда и расхождение в размере сети: 320 продаж дают не 320 школ, а около 480 открытий до закрытий.'));
children.push(p('Лаг открытия следующих школ решает многое:'));
children.push(table([0.2, 0.16, 0.16, 0.16, 0.16, 0.16], ['l', 'r', 'r', 'r', 'r', 'r'],
  ['Лаг, мес.', '0', '3 (сейчас)', '6', '12', '24'],
  [['EBITDA за 5 лет, млн ₽', '763', '731', '700', '638', '541']]));
children.push(p('Факт даёт только нижнюю границу: пять наблюдаемых вторых школ открылись через 0–3 месяца, но окно наблюдения — один сезон. Параметр остаётся на решении владельца модели.', { before: 8, italics: true, color: MUTED, size: 9 }));

children.push(h2('Постепенный набор детей'));
children.push(p('Принцип v4 сохранён: школа выходит на мощность этапами 1–12 месяцев, 13–24 месяца, зрелая (от 75 до 165 детей при полной загрузке). Заменён способ расчёта — вместо длинной формулы в помесячном листе используются когорты школ. Принцип роста не изменился.'));

children.push(h2('Тарифная матрица работает помесячно'));
children.push(p('Матрица «5 территорий × 4 формата» (паушальный от 390 тыс. до 1 490 тыс. ₽, дети от 12,5 до 250 на школу) была и в v4, но взвешенное среднее было статичным. Сейчас оно индексируется инфляцией помесячно вместе с ценами и расходами: паушальный 937 767 ₽ с НДС в первом сезоне, дальше по индексу.'));

children.push(h2('Кассовый разрыв закрывается правилом, а не деньгами'));
children.push(p('В v4 отрицательный остаток закрывался строкой «Привлечено внешнего финансирования» — 4,5 млн ₽ в базовом сценарии и 26,4 млн ₽ в негативном. Откуда эти деньги и по какой ставке, модель не говорила.'));
children.push(p('Сейчас такой строки нет: овердрафт владелец решил не заводить. Вместо неё действует ограничение — клубу возвращаем займ только тем, что остаётся сверх 10 млн ₽ на счёте и сверх резерва на ближайшую сезонную просадку. Результат: минимальный остаток за 60 месяцев +894 718 ₽, отрицательных месяцев нет, клубу возвращаются те же 50 млн ₽.'));
children.push(p('В негативном сценарии разрыв остаётся — девять месяцев, до −7,5 млн ₽. Там он операционный, и это открытый вопрос к владельцу.', { italics: true, color: MUTED, size: 9 }));

children.push(h2('БДР и БДДС разведены'));
children.push(p('В v4 начисление и деньги жили в одном помесячном листе. Сейчас это разные блоки, а расхождение между ними объясняется складом, займом и авансами. Сходимость проверяется формулой «мост от EBITDA к денежному потоку»: контроль должен давать ноль.'));

// 4
children.push(h1('4. Что добавилось, чего в v4 не было'));
children.push(bullet('Лагеря — участники, выручка, расходы, авансы и проживание. Сейчас выключены и ждут решения владельца.'));
children.push(bullet('Факт сезона 25/26 — помесячный P&L за сентябрь 2025 — июль 2026 и реестр начислений за октябрь 2025 — август 2026. Выручка 26,44 млн ₽, EBITDA −23,68 млн ₽. Включается в итог переключателем на дашборде.'));
children.push(bullet('Чувствительность — отдельный расчёт по шести драйверам.'));
children.push(bullet('Расчёт от любой даты старта: веса плана продаж и согласованный мерч ищутся по календарю, а не по номеру столбца.'));

// 5
children.push(h1('5. Результаты: было и стало'));
children.push(p('Базовый сценарий сегодня соответствует позитивному плану v4 — 320 продаж против 317, — поэтому даны обе колонки старого файла.'));
children.push(table([0.34, 0.22, 0.22, 0.22], ['l', 'r', 'r', 'r'],
  ['Показатель за 5 лет', 'v4 «Базовый»', 'v4 «Позитивный»', 'Сейчас «Базовый»'],
  [
    ['Продано франшиз, шт.', '210', '317', '320'],
    ['Активные школы на конец, шт.', '229', '338', '515'],
    ['Выручка УК без НДС, ₽', '1 537 338 920', '2 415 023 533', '2 181 352 193'],
    ['EBITDA, ₽', '382 900 614', '800 830 838', '731 487 728'],
    ['Рентабельность EBITDA', '24,9 %', '33,2 %', '33,5 %'],
    ['Деньги на конец, ₽', '299 662 194', '641 909 558', '498 613 411'],
    ['Внешнее финансирование, ₽', '4 474 785', '4 960 581', '0'],
    ['Погашение займа ФК, ₽', '50 000 000', '50 000 000', '50 000 000'],
  ], { boldLast: true }));
children.push(p('Школ больше при сопоставимом плане продаж — это 1,5 школы на партнёра. Выручка и EBITDA ниже позитивного сценария v4 при большем числе школ — это снижение среднего абонемента действующей сети с 7 800 до 6 000 ₽ и лаг открытия следующих школ.', { before: 10 }));
children.push(new Paragraph({
  spacing: { before: 140, after: 60, line: 260 },
  border: {
    top: { style: BorderStyle.SINGLE, size: 2, color: LINE, space: 6 },
    bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE, space: 6 },
  },
  children: [
    t('Оговорка. ', { bold: true, color: RED }),
    t('Метрику «дети на конец» из v4 сравнивать нельзя: в самом файле v4 базовый сценарий показывает 25 266 детей против 5 906 в позитивном — при меньшем числе школ и меньшей выручке. Это дефект расчёта старой модели, а не разница допущений.'),
  ],
}));

// 6
children.push(h1('6. Что осталось открытым'));
children.push(p('Решает владелец модели, не расчёт.'));
children.push(num([t('Лаг открытия следующих школ партнёра', { bold: true }), t(' — сейчас 3 месяца. Вилка EBITDA от 541 до 763 млн ₽.')]));
children.push(num([t('База выручки действующей сети занижена', { bold: true }), t(' — модель считает 272 млн ₽ за год, факт сезона 25/26 составил 417 млн ₽. Средний чек подтверждён, значит расходится число детей или профиль летней сезонности: индекс посещаемости падает до 6 %, а фактическая выручка июля — августа держится на уровне марта.')]));
children.push(num([t('Кассовый разрыв в негативном сценарии', { bold: true }), t(' — девять месяцев, до −7,5 млн ₽.')]));
children.push(num([t('Месяцы закупок склада', { bold: true }), t(' — квартальная закупка до цели одним платежом бьёт по кассе; помесячное пополнение сгладило бы её.')]));
children.push(num([t('Когда включать лагеря.', { bold: true })]));
children.push(num([t('P&L за август 2026', { bold: true }), t(' — единственный недостающий месяц факта.')]));

// ---- документ -----------------------------------------------------------
const doc = new Document({
  creator: 'Финмодель УК «Спартак»',
  title: 'Чем текущая модель отличается от v4.0',
  description: 'Сравнение финансовой модели франшизы Спартак v4.0 с текущей моделью Spartak_V9_0',
  styles: {
    default: {
      document: { run: { font: FONT, size: sz(10), color: INK } },
    },
  },
  numbering: {
    config: [
      { reference: 'dots', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 340, hanging: 200 } } } }] },
      { reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120 },
          children: [new TextRun({
            children: ['Финмодель УК «Спартак» · отличия от v4.0 · ', PageNumber.CURRENT, ' из ', PageNumber.TOTAL_PAGES],
            font: FONT, size: sz(8), color: MUTED,
          })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync(process.argv[2], b);
  console.log('готово:', process.argv[2], b.length, 'байт');
});
