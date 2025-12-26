const SHEET_NAMES = {
  PATRIMONIO_HIST: "PATRIMONIO_HIST",
  POSICOES: "POSICOES",
  BANCOS: "BANCOS",
  TIPOS: "TIPOS",
};

const POSICOES_HEADERS = [
  "date",
  "bank",
  "investment_type",
  "asset",
  "investment_date",
  "maturity_date",
  "invested_value",
  "current_value",
  "recurring",
  "recurring_amount",
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Dashboard Financeiro")
    .addItem("Abrir dashboard", "showDashboard")
    .addToUi();
  setupSheets();
}

function doGet() {
  setupSheets();
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Dashboard Financeiro")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showDashboard() {
  const html = HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Dashboard Financeiro")
    .setWidth(1200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  Object.values(SHEET_NAMES).forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      addHeaders(sheet, name);
    } else if (sheet.getLastRow() === 0) {
      addHeaders(sheet, name);
    } else {
      ensureHeaders(sheet, name);
    }
  });
}

function addHeaders(sheet, name) {
  const headers = {
    [SHEET_NAMES.PATRIMONIO_HIST]: ["date", "total_value", "notes"],
    [SHEET_NAMES.POSICOES]: POSICOES_HEADERS,
    [SHEET_NAMES.BANCOS]: ["bank"],
    [SHEET_NAMES.TIPOS]: ["investment_type"],
  }[name];

  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function ensureHeaders(sheet, name) {
  const headers = {
    [SHEET_NAMES.PATRIMONIO_HIST]: ["date", "total_value", "notes"],
    [SHEET_NAMES.POSICOES]: POSICOES_HEADERS,
    [SHEET_NAMES.BANCOS]: ["bank"],
    [SHEET_NAMES.TIPOS]: ["investment_type"],
  }[name];
  if (!headers) return;

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsUpdate = headers.some((h, idx) => existing[idx] !== h);
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getDashboardData(period, positionDate) {
  setupSheets();
  const patrimonio = readTable(SHEET_NAMES.PATRIMONIO_HIST, ["date", "total_value", "notes"]);
  const posicoes = readTable(SHEET_NAMES.POSICOES, POSICOES_HEADERS, true);
  const banks = readUniqueColumn(SHEET_NAMES.BANCOS, "bank");
  const types = readUniqueColumn(SHEET_NAMES.TIPOS, "investment_type");

  const alocacao = buildAlocacao(posicoes, positionDate);
  const syntheticSeries = buildSyntheticPatrimonioFromAlocacao(alocacao);
  const patrimonioData = buildPatrimonioData(patrimonio, period, syntheticSeries);

  return {
    patrimonioAtual: alocacao.totalCurrent,
    crescimentoMensalPercent: patrimonioData.crescimentoMensalPercent,
    crescimentoMensalValor: patrimonioData.crescimentoMensalValor,
    crescimento12MesesPercent: patrimonioData.crescimento12MesesPercent,
    seriesPatrimonio: patrimonioData.series,
    crescimentoMensalSeries: patrimonioData.crescimentoMensalSeries,
    alocacaoPorBanco: alocacao.porBanco,
    alocacaoPorTipo: alocacao.porTipo,
    resumo: {
      bancos: alocacao.resumoBancos,
      tipos: alocacao.resumoTipos,
    },
    datasPosicoes: alocacao.datasDisponiveis,
    lookups: { banks, types },
    totalInvestidoPosicoes: alocacao.totalInvested,
    totalAtualPosicoes: alocacao.totalCurrent,
    rentabilidadeTotalValor: alocacao.totalCurrent - alocacao.totalInvested,
    rentabilidadeTotalPercent: alocacao.totalInvested ? ((alocacao.totalCurrent - alocacao.totalInvested) / alocacao.totalInvested) * 100 : null,
    posicoesList: alocacao.posicoesList,
  };
}

function addPatrimonioEntry(entry) {
  setupSheets();
  const { date, total_value, notes } = entry;
  if (!date) throw new Error("Data é obrigatória.");
  const value = Number(total_value);
  if (isNaN(value)) throw new Error("Valor inválido.");

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.PATRIMONIO_HIST);
  sheet.appendRow([date, value, notes || ""]);
  return getDashboardData("12m");
}

function addPosicoesEntries(payload) {
  setupSheets();
  const { entries, newBanks, newTypes } = payload;
  if (!entries || !entries.length) throw new Error("Inclua ao menos uma posição.");
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.POSICOES);
  const rows = entries.map((item) => {
    if (!item.date || !item.bank || !item.investment_type) {
      throw new Error("Data, banco e tipo são obrigatórios.");
    }
    const invested = Number(item.invested_value);
    const current = Number(item.current_value);
    const recurringAmount = Number(item.recurring_amount || 0);
    if (isNaN(invested) || isNaN(current)) throw new Error("Valores inválidos.");
    return [
      item.date,
      item.bank,
      item.investment_type,
      item.asset || "",
      item.investment_date || "",
      item.maturity_date || "",
      invested,
      current,
      item.recurring === true || item.recurring === "true" || item.recurring === "on",
      isNaN(recurringAmount) ? 0 : recurringAmount,
    ];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, POSICOES_HEADERS.length).setValues(rows);

  if (Array.isArray(newBanks) && newBanks.length) {
    const bankSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.BANCOS);
    const uniqueBanks = filterNewUnique(bankSheet, newBanks);
    if (uniqueBanks.length) {
      bankSheet.getRange(bankSheet.getLastRow() + 1, 1, uniqueBanks.length, 1).setValues(uniqueBanks.map((b) => [b]));
    }
  }

  if (Array.isArray(newTypes) && newTypes.length) {
    const typeSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.TIPOS);
    const uniqueTypes = filterNewUnique(typeSheet, newTypes);
    if (uniqueTypes.length) {
      typeSheet.getRange(typeSheet.getLastRow() + 1, 1, uniqueTypes.length, 1).setValues(uniqueTypes.map((t) => [t]));
    }
  }

  return getDashboardData("12m");
}

function updatePosicaoEntry(entry) {
  setupSheets();
  const { rowNumber } = entry;
  if (!rowNumber) throw new Error("Linha inválida para edição.");
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.POSICOES);
  const invested = Number(entry.invested_value);
  const current = Number(entry.current_value);
  const recurringAmount = Number(entry.recurring_amount || 0);
  if (!entry.date || !entry.bank || !entry.investment_type) throw new Error("Data, banco e tipo são obrigatórios.");
  if (isNaN(invested) || isNaN(current)) throw new Error("Valores inválidos.");
  const values = [
    entry.date,
    entry.bank,
    entry.investment_type,
    entry.asset || "",
    entry.investment_date || "",
    entry.maturity_date || "",
    invested,
    current,
    entry.recurring === true || entry.recurring === "true" || entry.recurring === "on",
    isNaN(recurringAmount) ? 0 : recurringAmount,
  ];
  sheet.getRange(Number(rowNumber), 1, 1, POSICOES_HEADERS.length).setValues([values]);
  return getDashboardData("12m");
}

function readTable(sheetName, headers, includeRowNumber) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .map((row, idx) => {
      const obj = headers.reduce((acc, key, index) => {
        acc[key] = row[index];
        return acc;
      }, {});
      if (includeRowNumber) obj.rowNumber = idx + 2;
      return obj;
    })
    .filter((row) => Object.values(row).some((cell) => cell !== ""));
}

function readUniqueColumn(sheetName, key) {
  const data = readTable(sheetName, [key]);
  const set = new Set();
  data.forEach((item) => {
    if (item[key]) set.add(item[key]);
  });
  return Array.from(set);
}

function filterNewUnique(sheet, incoming) {
  if (!incoming || !incoming.length) return [];
  const last = sheet.getLastRow();
  if (last < 2) {
    return incoming.filter((item) => item);
  }
  const existing = new Set(
    sheet
      .getRange(2, 1, last - 1, 1)
      .getValues()
      .flat()
      .filter((v) => v !== "")
  );
  return incoming.filter((item) => item && !existing.has(item));
}

function buildPatrimonioData(records, period, fallbackSeries) {
  let parsed = records
    .map((item) => ({
      date: new Date(item.date),
      total_value: Number(item.total_value),
      notes: item.notes,
    }))
    .filter((item) => !isNaN(item.date.getTime()) && !isNaN(item.total_value))
    .sort((a, b) => a.date - b.date);

  if (!parsed.length && Array.isArray(fallbackSeries) && fallbackSeries.length) {
    parsed = fallbackSeries
      .map((item) => ({
        date: new Date(item.date),
        total_value: Number(item.total_value),
      }))
      .filter((item) => !isNaN(item.date.getTime()) && !isNaN(item.total_value))
      .sort((a, b) => a.date - b.date);
  }

  if (!parsed.length) {
    return {
      patrimonioAtual: 0,
      crescimentoMensalPercent: null,
      crescimentoMensalValor: null,
      crescimento12MesesPercent: null,
      series: [],
      crescimentoMensalSeries: [],
    };
  }

  const patrimonioAtual = parsed[parsed.length - 1].total_value;

  const monthlyClosing = buildMonthlyClosings(parsed);
  const crescimentoMensal = calculateMoM(monthlyClosing);
  const crescimento12MesesPercent = calculateYoY(monthlyClosing);

  const filteredSeries = filterByPeriod(parsed, period).map((item) => ({
    label: formatDateLabel(item.date),
    value: item.total_value,
  }));

  return {
    patrimonioAtual,
    crescimentoMensalPercent: crescimentoMensal.percent,
    crescimentoMensalValor: crescimentoMensal.delta,
    crescimento12MesesPercent,
    series: filteredSeries,
    crescimentoMensalSeries: monthlyClosing.map((item) => ({
      label: formatMonthLabel(item.date),
      value: item.percentGrowth,
    })),
  };
}

function buildMonthlyClosings(records) {
  const map = new Map();
  records.forEach((item) => {
    const key = `${item.date.getFullYear()}-${item.date.getMonth()}`;
    map.set(key, item);
  });

  const closings = Array.from(map.values()).sort((a, b) => a.date - b.date);
  for (let i = 1; i < closings.length; i++) {
    const prev = closings[i - 1].total_value;
    const curr = closings[i].total_value;
    closings[i].percentGrowth = prev ? ((curr - prev) / prev) * 100 : null;
  }
  return closings;
}

function calculateMoM(monthlyClosing) {
  if (monthlyClosing.length < 2) return { percent: null, delta: null };
  const last = monthlyClosing[monthlyClosing.length - 1];
  const prev = monthlyClosing[monthlyClosing.length - 2];
  const delta = last.total_value - prev.total_value;
  const percent = prev.total_value ? (delta / prev.total_value) * 100 : null;
  return { percent, delta };
}

function calculateYoY(monthlyClosing) {
  if (monthlyClosing.length < 13) return null;
  const last = monthlyClosing[monthlyClosing.length - 1];
  const targetDate = new Date(last.date);
  targetDate.setMonth(targetDate.getMonth() - 12);

  let candidate = null;
  monthlyClosing.forEach((item) => {
    if (item.date <= targetDate) candidate = item;
  });
  if (!candidate || !candidate.total_value) return null;
  return ((last.total_value - candidate.total_value) / candidate.total_value) * 100;
}

function filterByPeriod(records, period) {
  if (period === "all" || !period) return records;
  const months = { "3m": 3, "6m": 6, "12m": 12 }[period] || 12;
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);
  return records.filter((item) => item.date >= cutoff);
}

function formatDateLabel(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatMonthLabel(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM/yyyy");
}

function buildAlocacao(posicoes, positionDate) {
  if (!posicoes.length) {
    return {
      porBanco: [],
      porTipo: [],
      resumoBancos: [],
      resumoTipos: [],
      datasDisponiveis: [],
      totalCurrent: 0,
      totalInvested: 0,
    };
  }

  const parsed = posicoes
    .map((item) => ({
      date: new Date(item.date),
      bank: item.bank,
      investment_type: item.investment_type,
      asset: item.asset,
      investment_date: item.investment_date ? new Date(item.investment_date) : null,
      maturity_date: item.maturity_date ? new Date(item.maturity_date) : null,
      invested_value: Number(item.invested_value),
      current_value: Number(item.current_value),
      recurring: item.recurring === true || item.recurring === "true" || item.recurring === "on",
      recurring_amount: Number(item.recurring_amount || 0),
    }))
    .filter((item) => !isNaN(item.date.getTime()) && !isNaN(item.invested_value) && !isNaN(item.current_value))
    .sort((a, b) => a.date - b.date);

  const datasDisponiveis = Array.from(new Set(parsed.map((p) => formatDateLabel(p.date)))).sort();
  const targetDate = resolvePositionDate(parsed, positionDate);
  const filtered = parsed.filter((item) => formatDateLabel(item.date) === targetDate);
  const totalCurrent = filtered.reduce((sum, item) => sum + item.current_value, 0);
  const totalInvested = filtered.reduce((sum, item) => sum + item.invested_value, 0);

  const porBanco = aggregateByKey(filtered, "bank", totalCurrent);
  const porTipo = aggregateByKey(filtered, "investment_type", totalCurrent);

  const posicoesList = filtered.map((item) => ({
    ...item,
    date: formatDateLabel(item.date),
    investment_date: item.investment_date ? formatDateLabel(item.investment_date) : "",
    maturity_date: item.maturity_date ? formatDateLabel(item.maturity_date) : "",
    recurring_amount: Number(item.recurring_amount || 0),
  }));

  return {
    porBanco,
    porTipo,
    resumoBancos: porBanco,
    resumoTipos: porTipo,
    datasDisponiveis,
    totalCurrent,
    totalInvested,
    posicoesList,
  };
}

function buildSyntheticPatrimonioFromAlocacao(alocacao) {
  if (!alocacao.posicoesList || !alocacao.posicoesList.length) return [];
  const valid = alocacao.posicoesList
    .map((item) => ({
      investmentDate: item.investment_date ? new Date(item.investment_date) : null,
      currentDate: item.date ? new Date(item.date) : null,
      invested: Number(item.invested_value),
      current: Number(item.current_value),
    }))
    .filter((item) => item.investmentDate && item.currentDate && !isNaN(item.investmentDate.getTime()) && !isNaN(item.currentDate.getTime()) && item.invested > 0 && item.current >= 0);

  if (!valid.length) return [];

  const startDate = new Date(Math.min(...valid.map((v) => v.investmentDate.getTime())));
  const endDate = new Date(Math.max(...valid.map((v) => v.currentDate.getTime())));
  const totalInvested = valid.reduce((sum, v) => sum + v.invested, 0);
  const totalCurrent = valid.reduce((sum, v) => sum + v.current, 0);
  if (totalInvested <= 0) return [];

  const daysDiff = Math.max(1, Math.round((endDate - startDate) / (1000 * 60 * 60 * 24)));
  const periods = daysDiff || 1;
  const rate = Math.pow(totalCurrent / totalInvested || 1, 1 / periods) - 1;

  const series = [];
  for (let i = 0; i <= periods; i++) {
    const date = new Date(startDate);
    date.setDate(startDate.getDate() + i);
    const total_value = totalInvested * Math.pow(1 + rate, i);
    series.push({ date, total_value });
  }

  return series;
}

function resolvePositionDate(records, positionDate) {
  if (positionDate) return positionDate;
  const latest = records[records.length - 1];
  return formatDateLabel(latest.date);
}

function aggregateByKey(list, key, total) {
  const map = new Map();
  list.forEach((item) => {
    const current = map.get(item[key]) || { invested: 0, current: 0 };
    current.invested += item.invested_value;
    current.current += item.current_value;
    map.set(item[key], current);
  });
  return Array.from(map.entries())
    .map(([label, sums]) => {
      const profit = sums.current - sums.invested;
      return {
        label,
        invested: sums.invested,
        current: sums.current,
        value: sums.current,
        percent: total ? (sums.current / total) * 100 : 0,
        profit,
        profitPercent: sums.invested ? (profit / sums.invested) * 100 : null,
      };
    })
    .sort((a, b) => b.value - a.value);
}
