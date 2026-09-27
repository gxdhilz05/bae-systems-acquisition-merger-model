import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/workspace/scratch/d8279fcea226/bae-acquisition-merger-model/outputs/d8279fcea226";
const outputPath = `${outputDir}/bae_cohort_acquisition_model_v3.xlsx`;

const wb = Workbook.create();
const summary = wb.worksheets.add("Summary");
const assumptions = wb.worksheets.add("Assumptions");
const valuation = wb.worksheets.add("Valuation");
const purchaseAccounting = wb.worksheets.add("Purchase Accounting");
const screening = wb.worksheets.add("Screening");
const deal = wb.worksheets.add("Deal");
const earnings = wb.worksheets.add("Earnings");
const leverage = wb.worksheets.add("Leverage");
const sources = wb.worksheets.add("Sources & Guide");

const navy = "#17365D";
const blue = "#1F4E78";
const paleBlue = "#D9EAF7";
const paleGreen = "#E2F0D9";
const paleYellow = "#FFF2CC";
const paleRed = "#FCE4D6";
const grey = "#E7E6E6";
const dark = "#1F1F1F";
const inputBlue = "#0000FF";
const linkGreen = "#008000";
const fontFamily = "Arial";

function baseSheet(sheet) {
  sheet.showGridLines = false;
  sheet.getRange("A1:K45").format.font = { name: fontFamily, size: 10, color: dark };
  sheet.getRange("A1:K45").format.verticalAlignment = "center";
}

function title(sheet, text, subtitle) {
  sheet.getRange("C2").values = [[text]];
  sheet.getRange("C2").format.font = { name: fontFamily, size: 16, bold: true, color: navy };
  sheet.getRange("C3:J3").format.borders = { bottom: { style: "thin", color: navy } };
  if (subtitle) {
    sheet.getRange("C4").values = [[subtitle]];
    sheet.getRange("C4:J4").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };
  }
}

function section(sheet, range, text) {
  const band = sheet.getRange(range);
  band.getCell(0, 0).values = [[text]];
  band.format.fill = navy;
  band.format.font = { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" };
  band.format.borders = { preset: "outside", style: "thin", color: navy };
}

function header(sheet, range) {
  sheet.getRange(range).format.fill = blue;
  sheet.getRange(range).format.font = { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" };
  sheet.getRange(range).format.horizontalAlignment = "center";
  sheet.getRange(range).format.borders = { preset: "all", style: "thin", color: "#FFFFFF" };
}

function inputStyle(range) {
  range.format.fill = paleYellow;
  range.format.font = { name: fontFamily, size: 10, color: inputBlue };
  range.format.borders = { preset: "outside", style: "thin", color: "#D6B656" };
}

function crossLink(range) {
  range.format.font = { name: fontFamily, size: 10, color: linkGreen };
}

function totalStyle(range) {
  range.format.font = { name: fontFamily, size: 10, bold: true, color: dark };
  range.format.borders = { top: { style: "thin", color: dark }, bottom: { style: "double", color: dark } };
}

for (const sheet of [summary, assumptions, valuation, purchaseAccounting, screening, deal, earnings, leverage, sources]) baseSheet(sheet);

// Assumptions
title(assumptions, "BAE Systems acquisition of Cohort", "Editable assumptions are yellow with blue text. All amounts are £m unless stated.");
section(assumptions, "C7:D7", "Acquirer inputs");
assumptions.getRange("C8:D15").values = [
  ["Input", "Value"],
  ["BAE share price (p)", 1992.5],
  ["BAE underlying EPS (p)", 75.2],
  ["BAE diluted shares (m)", 2990.0],
  ["BAE underlying EBIT", 3322.0],
  ["BAE net debt", 3173.0],
  ["BAE cash and cash equivalents", 4200.0],
  ["BAE free cash flow", 2158.0],
];
header(assumptions, "C8:D8");
inputStyle(assumptions.getRange("D9:D15"));
assumptions.getRange("D9:D15").format.numberFormat = "#,##0.0;(#,##0.0);-";

section(assumptions, "F7:G7", "Target and market inputs");
assumptions.getRange("F8:G17").values = [
  ["Input", "Value"],
  ["Cohort share price (p)", 1140.0],
  ["Cohort diluted shares (m)", 46.03],
  ["Cohort adjusted operating profit", 36.3],
  ["Cohort adjusted EPS (p)", 61.9],
  ["Cohort net debt / (funds)", -2.2],
  ["Cohort order book", 618.8],
  ["GBP per USD", 0.7565],
  ["Market data date", new Date("2026-09-24T00:00:00Z")],
  ["Cohort book equity attributable to parent", 186.417],
];
header(assumptions, "F8:G8");
inputStyle(assumptions.getRange("G9:G17"));
assumptions.getRange("G9:G15").format.numberFormat = "#,##0.0;(#,##0.0);-";
assumptions.getRange("G15").format.numberFormat = "0.0000";
assumptions.getRange("G16").format.numberFormat = "dd-mmm-yy";
assumptions.getRange("G17").format.numberFormat = "#,##0.0;(#,##0.0);-";

section(assumptions, "C19:D19", "Illustrative transaction assumptions");
assumptions.getRange("C20:D29").values = [
  ["Assumption", "Base case"],
  ["Offer premium", 0.30],
  ["Debt funding share", 0.60],
  ["Cash funding share", null],
  ["New debt interest rate", 0.05],
  ["Foregone cash interest rate", 0.035],
  ["Tax rate", 0.25],
  ["Run-rate pre-tax synergies", 12.0],
  ["One-time integration costs", 18.0],
  ["Transaction fees", 15.0],
];
assumptions.getRange("D23").formulas = [["=1-D22"]];
header(assumptions, "C20:D20");
inputStyle(assumptions.getRange("D21:D22"));
inputStyle(assumptions.getRange("D24:D29"));
assumptions.getRange("D21:D26").format.numberFormat = "0.0%";
assumptions.getRange("D27:D29").format.numberFormat = "#,##0.0;(#,##0.0);-";

section(assumptions, "F19:G19", "Simple valuation assumptions");
assumptions.getRange("F20:G30").values = [
  ["Assumption", "Base case"],
  ["FY2026 revenue", 306.4],
  ["FY2027 revenue growth", 0.08],
  ["Annual growth fade", 0.01],
  ["FY2027 operating margin", 0.125],
  ["Annual margin increase", 0.005],
  ["D&A as % of revenue", 0.025],
  ["Capital expenditure as % of revenue", 0.04],
  ["Working capital as % of revenue increase", 0.10],
  ["WACC", 0.09],
  ["Terminal growth", 0.025],
];
header(assumptions, "F20:G20");
inputStyle(assumptions.getRange("G21:G30"));
assumptions.getRange("G21").format.numberFormat = "#,##0.0;(#,##0.0);-";
assumptions.getRange("G22:G30").format.numberFormat = "0.0%";

section(assumptions, "C37:D37", "Purchase accounting assumptions");
assumptions.getRange("C38:D40").values = [
  ["Assumption", "Base case"],
  ["Premium allocated to new intangibles", 0.35],
  ["New intangible useful life (years)", 10],
];
header(assumptions, "C38:D38");
inputStyle(assumptions.getRange("D39:D40"));
assumptions.getRange("D39").format.numberFormat = "0.0%";
assumptions.getRange("D40").format.numberFormat = "0";
assumptions.getRange("C32:H34").values = [["How to use this tab", null, null, null, null, null], ["Change only the yellow blue-text cells. The offer, earnings and leverage schedules update automatically.", null, null, null, null, null], ["Base-case inputs are illustrative, not management guidance or a formal offer recommendation.", null, null, null, null, null]];
assumptions.getRange("C32:H32").format.font = { name: fontFamily, size: 10, bold: true, color: navy };
assumptions.getRange("C33:H34").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };

// Simple Cohort forecast and valuation
title(valuation, "Cohort standalone valuation", "A five-year DCF with a simple synergy-supported price ceiling. All amounts are £m unless stated.");
section(valuation, "C7:I7", "Forecast and unlevered free cash flow");
valuation.getRange("C8:I20").values = [
  ["Metric", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"],
  ["Revenue", null, null, null, null, null, null],
  ["Revenue growth", null, null, null, null, null, null],
  ["Adjusted operating margin", null, null, null, null, null, null],
  ["Adjusted operating profit", null, null, null, null, null, null],
  ["Cash tax", null, null, null, null, null, null],
  ["NOPAT", null, null, null, null, null, null],
  ["D&A", null, null, null, null, null, null],
  ["Capital expenditure", null, null, null, null, null, null],
  ["Working capital investment", null, null, null, null, null, null],
  ["Unlevered free cash flow", null, null, null, null, null, null],
  ["Discount factor", null, null, null, null, null, null],
  ["Present value of UFCF", null, null, null, null, null, null],
];
header(valuation, "C8:I8");
valuation.getRange("D9:D12").formulas = [
  ["='Assumptions'!G21"],
  ["=D9/270-1"],
  ["='Assumptions'!G11/'Assumptions'!G21"],
  ["='Assumptions'!G11"],
];
valuation.getRange("E10").formulas = [["='Assumptions'!G22"]];
valuation.getRange("F10:I10").formulasR1C1 = [["=MAX(RC[-1]-'Assumptions'!R23C7,0)", "=MAX(RC[-1]-'Assumptions'!R23C7,0)", "=MAX(RC[-1]-'Assumptions'!R23C7,0)", "=MAX(RC[-1]-'Assumptions'!R23C7,0)"]];
valuation.getRange("E9").formulas = [["=D9*(1+E10)"]];
valuation.getRange("E9:I9").fillRight();
valuation.getRange("E11").formulas = [["='Assumptions'!G24"]];
valuation.getRange("F11:I11").formulasR1C1 = [["=MIN(RC[-1]+'Assumptions'!R25C7,15%)", "=MIN(RC[-1]+'Assumptions'!R25C7,15%)", "=MIN(RC[-1]+'Assumptions'!R25C7,15%)", "=MIN(RC[-1]+'Assumptions'!R25C7,15%)"]];
valuation.getRange("E12").formulas = [["=E9*E11"]];
valuation.getRange("E12:I12").fillRight();
valuation.getRange("E13").formulas = [["=-E12*'Assumptions'!$D$26"]];
valuation.getRange("E13:I13").fillRight();
valuation.getRange("E14").formulas = [["=E12+E13"]];
valuation.getRange("E14:I14").fillRight();
valuation.getRange("E15").formulas = [["=E9*'Assumptions'!$G$26"]];
valuation.getRange("E15:I15").fillRight();
valuation.getRange("E16").formulas = [["=-E9*'Assumptions'!$G$27"]];
valuation.getRange("E16:I16").fillRight();
valuation.getRange("E17").formulas = [["=-(E9-D9)*'Assumptions'!$G$28"]];
valuation.getRange("E17:I17").fillRight();
valuation.getRange("E18").formulas = [["=SUM(E14:E17)"]];
valuation.getRange("E18:I18").fillRight();
valuation.getRange("E19").formulas = [["=1/(1+'Assumptions'!$G$29)^(COLUMN()-COLUMN($D$19))"]];
valuation.getRange("E19:I19").fillRight();
valuation.getRange("E20").formulas = [["=E18*E19"]];
valuation.getRange("E20:I20").fillRight();
crossLink(valuation.getRange("D9:D12"));
crossLink(valuation.getRange("E10:E11"));
valuation.getRange("D9:I9").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
valuation.getRange("D10:I11").format.numberFormat = "0.0%";
valuation.getRange("D12:I18").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
valuation.getRange("D19:I19").format.numberFormat = "0.000x";
valuation.getRange("D20:I20").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
totalStyle(valuation.getRange("C18:I18"));

section(valuation, "C24:D24", "Standalone DCF");
valuation.getRange("C25:D37").values = [
  ["Metric", "Value"],
  ["Present value of forecast UFCF", null],
  ["Terminal-year UFCF", null],
  ["Terminal value", null],
  ["Present value of terminal value", null],
  ["Enterprise value", null],
  ["Net funds", null],
  ["Equity value", null],
  ["Diluted shares (m)", null],
  ["DCF value per share (p)", null],
  ["Current share price (p)", null],
  ["Proposed offer price (p)", null],
  ["Offer premium to DCF", null],
];
header(valuation, "C25:D25");
valuation.getRange("D26:D37").formulas = [
  ["=SUM(E20:I20)"],
  ["=I18*(1+'Assumptions'!G30)"],
  ["=D27/('Assumptions'!G29-'Assumptions'!G30)"],
  ["=D28*I19"],
  ["=D26+D29"],
  ["=-'Assumptions'!G13"],
  ["=D30+D31"],
  ["='Assumptions'!G10"],
  ["=D32/D33*100"],
  ["='Assumptions'!G9"],
  ["='Deal'!D11"],
  ["=D36/D34-1"],
];
crossLink(valuation.getRange("D31:D33"));
crossLink(valuation.getRange("D35:D36"));
valuation.getRange("D26:D33").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
valuation.getRange("D34:D36").format.numberFormat = "#,##0.0p";
valuation.getRange("D37").format.numberFormat = "0.0%";
totalStyle(valuation.getRange("C30:D30"));
totalStyle(valuation.getRange("C32:D34"));

section(valuation, "F24:G24", "Price discipline");
valuation.getRange("F25:G34").values = [
  ["Metric", "Value"],
  ["Run-rate pre-tax synergies", null],
  ["After-tax synergies", null],
  ["Present value of synergies", null],
  ["Maximum supported equity value", null],
  ["Maximum supported price (p)", null],
  ["Maximum supported premium", null],
  ["Proposed premium", null],
  ["Offer above supported price (p)", null],
  ["Conclusion", "30% premium is not supported by the base case"],
];
header(valuation, "F25:G25");
valuation.getRange("G26:G33").formulas = [
  ["='Assumptions'!D27"],
  ["=G26*(1-'Assumptions'!D26)"],
  ["=G27/'Assumptions'!G29"],
  ["=D32+G28"],
  ["=G29/D33*100"],
  ["=G30/D35-1"],
  ["='Assumptions'!D21"],
  ["=D36-G30"],
];
crossLink(valuation.getRange("G26"));
crossLink(valuation.getRange("G32"));
valuation.getRange("G26:G29").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
valuation.getRange("G30:G31").format.numberFormat = "#,##0.0p";
valuation.getRange("G31:G32").format.numberFormat = "0.0%";
valuation.getRange("G33").format.numberFormat = "#,##0.0p";
valuation.getRange("G34").format.fill = paleRed;
valuation.getRange("G34").format.font = { name: fontFamily, size: 10, bold: true, color: "#9C0006" };
totalStyle(valuation.getRange("F29:G30"));

// Simplified purchase accounting and offer-price sensitivity
title(purchaseAccounting, "Simplified purchase accounting", "Shows how the purchase premium becomes intangible assets, deferred tax and goodwill. All amounts are £m unless stated.");
section(purchaseAccounting, "C7:D7", "Purchase price allocation");
purchaseAccounting.getRange("C8:D18").values = [
  ["Metric", "Value"],
  ["Equity purchase price", null],
  ["Book equity attributable to Cohort shareholders", null],
  ["Premium over book equity", null],
  ["New identifiable intangible assets", null],
  ["Deferred tax liability on write-up", null],
  ["New goodwill", null],
  ["Allocation check", null],
  ["New intangible useful life (years)", null],
  ["Annual amortisation, pre-tax", null],
  ["Annual amortisation, after-tax", null],
];
header(purchaseAccounting, "C8:D8");
purchaseAccounting.getRange("D9:D18").formulas = [
  ["='Deal'!D13"],
  ["='Assumptions'!G17"],
  ["=D9-D10"],
  ["=D11*'Assumptions'!D39"],
  ["=D12*'Assumptions'!D26"],
  ["=D11-D12+D13"],
  ["=D14+D12-D13-D11"],
  ["='Assumptions'!D40"],
  ["=D12/D16"],
  ["=D17*(1-'Assumptions'!D26)"],
];
crossLink(purchaseAccounting.getRange("D9:D10"));
crossLink(purchaseAccounting.getRange("D16"));
purchaseAccounting.getRange("D9:D15").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
purchaseAccounting.getRange("D16").format.numberFormat = "0";
purchaseAccounting.getRange("D17:D18").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
totalStyle(purchaseAccounting.getRange("C11:D11"));
totalStyle(purchaseAccounting.getRange("C14:D14"));
totalStyle(purchaseAccounting.getRange("C17:D18"));
purchaseAccounting.getRange("D15").conditionalFormats.add("cellIs", { operator: "notEqual", formula: 0, format: { fill: paleRed, font: { bold: true, color: "#9C0006" } } });

section(purchaseAccounting, "F7:K7", "Offer-price sensitivity");
purchaseAccounting.getRange("F8:K13").values = [
  ["Offer premium", "Offer price (p)", "Equity purchase price", "New intangibles", "EPS accretion after PPA", "Price vs supported ceiling (p)"],
  [0.10, null, null, null, null, null],
  [0.15, null, null, null, null, null],
  [0.20, null, null, null, null, null],
  [0.25, null, null, null, null, null],
  [0.30, null, null, null, null, null],
];
header(purchaseAccounting, "F8:K8");
purchaseAccounting.getRange("G9").formulas = [["='Assumptions'!$G$9*(1+F9)"]];
purchaseAccounting.getRange("G9:G13").fillDown();
purchaseAccounting.getRange("H9").formulas = [["=G9*'Assumptions'!$G$10/100"]];
purchaseAccounting.getRange("H9:H13").fillDown();
purchaseAccounting.getRange("I9").formulas = [["=MAX(H9-'Assumptions'!$G$17,0)*'Assumptions'!$D$39"]];
purchaseAccounting.getRange("I9:I13").fillDown();
purchaseAccounting.getRange("J9").formulas = [["=((('Assumptions'!$D$10*'Assumptions'!$D$11/100)+('Assumptions'!$G$12*'Assumptions'!$G$10/100)+('Assumptions'!$D$27*(1-'Assumptions'!$D$26))-((H9+'Assumptions'!$G$13+'Assumptions'!$D$29)*('Assumptions'!$D$22*'Assumptions'!$D$24+'Assumptions'!$D$23*'Assumptions'!$D$25)*(1-'Assumptions'!$D$26))-(I9/'Assumptions'!$D$40*(1-'Assumptions'!$D$26)))/'Assumptions'!$D$11*100)/'Assumptions'!$D$10-1"]];
purchaseAccounting.getRange("J9:J13").fillDown();
purchaseAccounting.getRange("K9").formulas = [["=G9-'Valuation'!$G$30"]];
purchaseAccounting.getRange("K9:K13").fillDown();
purchaseAccounting.getRange("F9:F13").format.numberFormat = "0.0%";
purchaseAccounting.getRange("G9:G13").format.numberFormat = "#,##0.0p";
purchaseAccounting.getRange("H9:I13").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
purchaseAccounting.getRange("J9:J13").format.numberFormat = "0.0%";
purchaseAccounting.getRange("K9:K13").format.numberFormat = "#,##0.0p;[Red](#,##0.0p);-";
purchaseAccounting.getRange("F13:K13").format.fill = paleRed;

section(purchaseAccounting, "C22:H22", "Interview interpretation");
purchaseAccounting.getRange("C23:H27").values = [
  ["Purchase accounting changes reported earnings but does not itself use additional cash.", null, null, null, null, null],
  ["New identifiable intangibles are amortised; goodwill is tested for impairment rather than amortised.", null, null, null, null, null],
  ["The deferred tax liability reflects the assumed difference between accounting and tax values.", null, null, null, null, null],
  ["At a 30% premium, purchase-accounting amortisation reduces run-rate EPS accretion to approximately 0.1%.", null, null, null, null, null],
  ["Including integration costs, the transaction is approximately 0.5% EPS dilutive in Year 1.", null, null, null, null, null],
];
purchaseAccounting.getRange("C23:H27").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };

// Screening
title(screening, "Target screening", "Scores organise judgment; they do not replace the evidence in the research files.");
screening.getRange("C7:J7").values = [["Category", "Weight", "Chemring", "Cohort", "Avon", "Chemring points", "Cohort points", "Avon points"]];
header(screening, "C7:J7");
screening.getRange("C8:F13").values = [
  ["Strategic and product fit", 0.25, 4, 5],
  ["Growth and financial quality", 0.20, 4, 4],
  ["Valuation and affordability", 0.20, 2, 4],
  ["Synergy potential", 0.15, 4, 4],
  ["Financing and leverage impact", 0.10, 4, 5],
  ["Execution and regulatory risk", 0.10, 3, 3],
];
screening.getRange("G8:G13").values = [[3],[4],[4],[3],[5],[4]];
screening.getRange("H8").formulas = [["=E8/5*$D8"]];
screening.getRange("H8:H13").fillDown();
screening.getRange("I8").formulas = [["=F8/5*$D8"]];
screening.getRange("I8:I13").fillDown();
screening.getRange("J8").formulas = [["=G8/5*$D8"]];
screening.getRange("J8:J13").fillDown();
screening.getRange("C14:J14").values = [["Weighted score", null, null, null, null, null, null, null]];
screening.getRange("D14").formulas = [["=SUM(D8:D13)"]];
screening.getRange("H14").formulas = [["=SUM(H8:H13)"]];
screening.getRange("I14").formulas = [["=SUM(I8:I13)"]];
screening.getRange("J14").formulas = [["=SUM(J8:J13)"]];
screening.getRange("E14").formulas = [["=H14*100"]];
screening.getRange("F14").formulas = [["=I14*100"]];
screening.getRange("G14").formulas = [["=J14*100"]];
screening.getRange("D8:D14").format.numberFormat = "0%";
screening.getRange("H8:J14").format.numberFormat = "0.0%";
screening.getRange("E14:G14").format.numberFormat = "0";
totalStyle(screening.getRange("C14:J14"));
screening.getRange("F8:F13").format.fill = paleGreen;

section(screening, "C18:J18", "Market reference at 24 September 2026");
screening.getRange("C19:J22").values = [
  ["Company", "Share price (p)", "Equity value", "Net debt / (funds)", "Enterprise value", "Adjusted operating profit", "EV / operating profit", "Observation"],
  ["Chemring", 535.0, 1440.0, 144.5, null, 73.5, null, "Strong fit; high valuation and capex burden"],
  ["Cohort", 1140.0, 525.0, -2.2, null, 36.3, null, "Best balance of fit, quality and affordability"],
  ["Avon Technologies", 1800.0, 530.0, 43.9, null, 30.5, null, "Affordable; strategically further from BAE core"],
];
header(screening, "C19:J19");
screening.getRange("G20").formulas = [["=E20+F20"]];
screening.getRange("G20:G22").fillDown();
screening.getRange("I20").formulas = [["=G20/H20"]];
screening.getRange("I20:I22").fillDown();
screening.getRange("D20:D22").format.numberFormat = "#,##0.0";
screening.getRange("E20:H22").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
screening.getRange("I20:I22").format.numberFormat = "0.0x";
screening.getRange("C26:J29").values = [
  ["Selection", "Cohort", null, null, null, null, null, null],
  ["Why", "Naval sensors, communications, fire control and electronic warfare fit BAE platforms and bids.", null, null, null, null, null, null],
  ["Main condition", "Preserve subsidiary autonomy and customer neutrality.", null, null, null, null, null, null],
  ["Main uncertainty", "Cash conversion and realisable revenue synergies require diligence.", null, null, null, null, null, null],
];
screening.getRange("C26:C29").format.font = { name: fontFamily, size: 10, bold: true, color: navy };
screening.getRange("D26:J29").format.wrapText = false;

// Deal schedule
title(deal, "Offer and sources & uses", "The offer premium and funding mix are editable on Assumptions.");
section(deal, "C7:D7", "Offer calculation");
deal.getRange("C8:D15").values = [["Metric", "Value"], ["Unaffected share price (p)", null], ["Offer premium", null], ["Offer price per share (p)", null], ["Diluted shares (m)", null], ["Equity purchase price", null], ["Target net debt / (funds)", null], ["Implied enterprise value", null]];
header(deal, "C8:D8");
deal.getRange("D9:D10").formulas = [["='Assumptions'!G9"],["='Assumptions'!D21"]];
deal.getRange("D11").formulas = [["=D9*(1+D10)"]];
deal.getRange("D12").formulas = [["='Assumptions'!G10"]];
deal.getRange("D13").formulas = [["=D11*D12/100"]];
deal.getRange("D14").formulas = [["='Assumptions'!G13"]];
deal.getRange("D15").formulas = [["=D13+D14"]];
crossLink(deal.getRange("D9:D10"));
crossLink(deal.getRange("D12"));
crossLink(deal.getRange("D14"));
deal.getRange("D9:D12").format.numberFormat = "#,##0.0;(#,##0.0);-";
deal.getRange("D10").format.numberFormat = "0.0%";
deal.getRange("D13:D15").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
totalStyle(deal.getRange("C15:D15"));

section(deal, "F7:G7", "Sources and uses");
deal.getRange("F8:G13").values = [["Uses", "Value"], ["Equity purchase price", null], ["Target net debt / (funds)", null], ["Transaction fees", null], ["Total transaction uses", null], ["One-time integration costs (separate)", null]];
header(deal, "F8:G8");
deal.getRange("G9:G12").formulas = [["=D13"],["=D14"],["='Assumptions'!D29"],["=SUM(G9:G11)"]];
deal.getRange("G13").formulas = [["='Assumptions'!D28"]];
crossLink(deal.getRange("G11"));
crossLink(deal.getRange("G13"));
deal.getRange("G9:G13").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
totalStyle(deal.getRange("F12:G12"));

deal.getRange("F16:G20").values = [["Sources", "Value"], ["New debt", null], ["BAE cash", null], ["Total sources", null], ["Funding check", null]];
header(deal, "F16:G16");
deal.getRange("G17:G20").formulas = [["=G12*'Assumptions'!D22"],["=G12*'Assumptions'!D23"],["=SUM(G17:G18)"],["=G19-G12"]];
crossLink(deal.getRange("G17:G18"));
deal.getRange("G17:G20").format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
totalStyle(deal.getRange("F19:G19"));
deal.getRange("G20").conditionalFormats.add("cellIs", { operator: "notEqual", formula: 0, format: { fill: paleRed, font: { bold: true, color: "#9C0006" } } });

section(deal, "C19:D19", "Valuation reference");
deal.getRange("C20:D24").values = [["Metric", "Value"], ["Current EV / adjusted operating profit", null], ["Offer EV / adjusted operating profit", null], ["Premium paid to current EV", null], ["Offer EV / order book", null]];
header(deal, "C20:D20");
deal.getRange("D21:D24").formulas = [["=('Assumptions'!G9*'Assumptions'!G10/100+'Assumptions'!G13)/'Assumptions'!G11"],["=D15/'Assumptions'!G11"],["=D15/('Assumptions'!G9*'Assumptions'!G10/100+'Assumptions'!G13)-1"],["=D15/'Assumptions'!G14"]];
crossLink(deal.getRange("D21:D24"));
deal.getRange("D21:D22").format.numberFormat = "0.0x";
deal.getRange("D23").format.numberFormat = "0.0%";
deal.getRange("D24").format.numberFormat = "0.0x";

// Earnings
title(earnings, "Illustrative earnings impact", "Shows EPS before and after simplified purchase-accounting amortisation.");
section(earnings, "C7:D7", "Run-rate earnings build");
earnings.getRange("C8:D26").values = [
  ["Metric", "Value"],
  ["BAE underlying EPS (p)", null],
  ["BAE diluted shares (m)", null],
  ["BAE standalone underlying earnings", null],
  ["Cohort adjusted net income", null],
  ["Run-rate synergies, pre-tax", null],
  ["New debt interest, pre-tax", null],
  ["Foregone cash interest, pre-tax", null],
  ["Net transaction effect, after tax", null],
  ["Pro forma underlying earnings", null],
  ["Pro forma diluted shares (m)", null],
  ["Pro forma EPS (p)", null],
  ["Run-rate EPS accretion / (dilution)", null],
  ["EPS accretion / (dilution) before synergies", null],
  ["Year 1 EPS impact incl. integration costs", null],
  ["New intangible amortisation, pre-tax", null],
  ["Pro forma earnings after purchase accounting", null],
  ["Run-rate EPS accretion after purchase accounting", null],
  ["Year 1 EPS impact after purchase accounting", null],
];
header(earnings, "C8:D8");
earnings.getRange("D9:D10").formulas = [["='Assumptions'!D10"],["='Assumptions'!D11"]];
earnings.getRange("D11").formulas = [["=D9*D10/100"]];
earnings.getRange("D12").formulas = [["='Assumptions'!G12*'Assumptions'!G10/100"]];
earnings.getRange("D13").formulas = [["='Assumptions'!D27"]];
earnings.getRange("D14").formulas = [["='Deal'!G17*'Assumptions'!D24"]];
earnings.getRange("D15").formulas = [["='Deal'!G18*'Assumptions'!D25"]];
earnings.getRange("D16").formulas = [["=D12+(D13-D14-D15)*(1-'Assumptions'!D26)"]];
earnings.getRange("D17").formulas = [["=D11+D16"]];
earnings.getRange("D18").formulas = [["=D10"]];
earnings.getRange("D19").formulas = [["=D17/D18*100"]];
earnings.getRange("D20").formulas = [["=D19/D9-1"]];
earnings.getRange("D21").formulas = [["=(D11+D12+(-D14-D15)*(1-'Assumptions'!D26))/D18*100/D9-1"]];
earnings.getRange("D22").formulas = [["=(D17-'Assumptions'!D28*(1-'Assumptions'!D26))/D18*100/D9-1"]];
earnings.getRange("D23").formulas = [["='Purchase Accounting'!D17"]];
earnings.getRange("D24").formulas = [["=D17-'Purchase Accounting'!D18"]];
earnings.getRange("D25").formulas = [["=(D24/D18*100)/D9-1"]];
earnings.getRange("D26").formulas = [["=((D24-'Assumptions'!D28*(1-'Assumptions'!D26))/D18*100)/D9-1"]];
crossLink(earnings.getRange("D9:D10"));
crossLink(earnings.getRange("D12:D15"));
crossLink(earnings.getRange("D23"));
earnings.getRange("D9:D10").format.numberFormat = "#,##0.0";
earnings.getRange("D11:D18").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
earnings.getRange("D19").format.numberFormat = "0.0p";
earnings.getRange("D20:D22").format.numberFormat = "0.0%";
earnings.getRange("D23:D24").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
earnings.getRange("D25:D26").format.numberFormat = "0.0%";
totalStyle(earnings.getRange("C19:D20"));
totalStyle(earnings.getRange("C24:D26"));
earnings.getRange("F9:H14").values = [
  ["Interpretation", null, null],
  ["Positive accretion is not proof of value creation.", null, null],
  ["A premium can make EPS accretive while still earning a weak return.", null, null],
  ["Purchase-accounting amortisation reduces reported EPS but is non-cash.", null, null],
  ["Revenue synergies are excluded from the base case.", null, null],
  ["The financing and synergy assumptions require diligence.", null, null],
];
earnings.getRange("F9:H9").format.font = { name: fontFamily, size: 10, bold: true, color: navy };
earnings.getRange("F10:H14").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };

// Leverage
title(leverage, "Illustrative leverage impact", "Net debt / EBIT is a screening proxy because a complete EBITDA schedule has not yet been built.");
section(leverage, "C7:D7", "Pro forma leverage build");
leverage.getRange("C8:D19").values = [
  ["Metric", "Value"],
  ["BAE standalone net debt", null],
  ["New debt funding", null],
  ["Cash funding", null],
  ["Target net debt / (funds)", null],
  ["Pro forma net debt", null],
  ["BAE underlying EBIT", null],
  ["Cohort adjusted operating profit", null],
  ["Run-rate synergies", null],
  ["Pro forma EBIT proxy", null],
  ["Standalone net debt / EBIT", null],
  ["Pro forma net debt / EBIT", null],
];
header(leverage, "C8:D8");
leverage.getRange("D9:D12").formulas = [["='Assumptions'!D13"],["='Deal'!G17"],["='Deal'!G18"],["='Assumptions'!G13"]];
leverage.getRange("D13").formulas = [["=SUM(D9:D12)"]];
leverage.getRange("D14:D16").formulas = [["='Assumptions'!D12"],["='Assumptions'!G11"],["='Assumptions'!D27"]];
leverage.getRange("D17").formulas = [["=SUM(D14:D16)"]];
leverage.getRange("D18").formulas = [["=D9/D14"]];
leverage.getRange("D19").formulas = [["=D13/D17"]];
crossLink(leverage.getRange("D9:D12"));
crossLink(leverage.getRange("D14:D16"));
leverage.getRange("D9:D17").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
leverage.getRange("D18:D19").format.numberFormat = "0.00x";
totalStyle(leverage.getRange("C13:D13"));
totalStyle(leverage.getRange("C17:D17"));
leverage.getRange("F9:H13").values = [["Reading the result", null, null], ["The transaction increases net debt by the full uses funded.", null, null], ["Target net funds reduce the enterprise value acquired.", null, null], ["The ratio remains illustrative until EBITDA, pensions and lease treatment are aligned.", null, null], ["Investment-grade impact requires rating-agency methodology and liquidity analysis.", null, null]];
leverage.getRange("F9:H9").format.font = { name: fontFamily, size: 10, bold: true, color: navy };
leverage.getRange("F10:H13").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };

// Summary
title(summary, "BAE Systems acquisition of Cohort", "Illustrative screening and acquisition model as at 24 September 2026");
summary.getRange("C7:D7").values = [["Headline", "Value"]];
header(summary, "C7:D7");
summary.getRange("C8:D18").values = [["Preferred target", "Cohort"], ["Cohort screening score", null], ["Illustrative offer price (p)", null], ["Illustrative equity purchase price", null], ["Illustrative enterprise value", null], ["Run-rate EPS accretion after purchase accounting", null], ["Year 1 EPS impact after purchase accounting", null], ["Pro forma net debt / EBIT", null], ["Standalone DCF value (p)", null], ["Maximum supported price incl. synergies (p)", null], ["Recommendation", "Pursue only below the supported price"]];
summary.getRange("D9:D17").formulas = [["='Screening'!F14"],["='Deal'!D11"],["='Deal'!D13"],["='Deal'!D15"],["='Earnings'!D25"],["='Earnings'!D26"],["='Leverage'!D19"],["='Valuation'!D34"],["='Valuation'!G30"]];
summary.getRange("D9").format.numberFormat = "0";
summary.getRange("D10").format.numberFormat = "#,##0.0p";
summary.getRange("D11:D12").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
summary.getRange("D13:D14").format.numberFormat = "0.0%";
summary.getRange("D15").format.numberFormat = "0.00x";
summary.getRange("D16:D17").format.numberFormat = "#,##0.0p";
summary.getRange("D8:D18").format.fill = paleBlue;
summary.getRange("C8:C18").format.font = { name: fontFamily, size: 10, bold: true, color: dark };
summary.getRange("F8:H15").values = [
  ["Current conclusion", null, null],
  ["Cohort provides the clearest link between specialist products and BAE platforms.", null, null],
  ["The proposed ownership model should preserve subsidiary autonomy.", null, null],
  ["The base case includes cost synergies but no revenue synergies.", null, null],
  ["The standalone DCF is below the proposed 30% premium offer.", null, null],
  ["Purchase accounting reduces run-rate EPS accretion to approximately 0.1%.", null, null],
  ["Year 1 is approximately 0.5% dilutive after integration costs and amortisation.", null, null],
  ["Recommendation: pursue Cohort only below approximately 1,300p per share.", null, null],
];
summary.getRange("F8:H8").format.fill = navy;
summary.getRange("F8:H8").format.font = { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" };
summary.getRange("F9:H15").format.wrapText = true;
summary.getRange("F9:H15").format.fill = "#F7F9FC";

section(summary, "C19:E19", "Decision conditions");
summary.getRange("C20:E24").values = [
  ["Condition", "Why it matters", "Current treatment"],
  ["Price discipline", "A high premium can absorb strategic benefits.", "30% premium is illustrative."],
  ["Autonomy", "Cohort's operating model supports agility and retention.", "Light-touch ownership assumed."],
  ["Customer neutrality", "Other primes may reduce purchases from a BAE-owned supplier.", "No revenue synergy included."],
  ["Cash conversion", "FY2026 receivables reduced operating cash generation.", "Forecast uses normalised working capital."],
];
header(summary, "C20:E20");
summary.getRange("C21:E24").format.wrapText = true;

// Sources and guide
title(sources, "Sources and model guide", "Reported information is separated from illustrative assumptions.");
section(sources, "C7:F7", "Key sources");
sources.getRange("C8:J16").values = [
  ["ID", "Company", "Document", "Date", "Model use", "URL", null, null],
  ["BAE-001", "BAE Systems", "2026 Half-yearly Report", "30-Jul-2026", "Cash, net debt and current trading", "https://investors.baesystems.com/dam/jcr:f84dc82f-4530-419c-b993-74ebb562bf97/HY-2026-Half-Yearly-Report-Final.04078ad460b3959f428f65bc388abb6f.pdf", null, null],
  ["BAE-003", "BAE Systems", "Annual Report 2025", "18-Feb-2026", "FY2025 EBIT, EPS and free cash flow", "https://investors.baesystems.com/dam/jcr%3A105fe9f2-cff7-4960-9d99-956aba996540/BAE-Systems-Annual-Report-2025.pdf", null, null],
  ["COH-001", "Cohort", "Annual Report and Accounts 2026", "19-Aug-2026", "Financials, divisions, order book and risks", "https://www.cohortplc.com/download_file/force/815/241", null, null],
  ["CHG-001", "Chemring", "FY2026 Interim Results", "02-Jun-2026", "Latest Chemring financials", "https://www.chemring.com/~/media/Files/C/Chemring-V3/press-releases/interim-results-2026.pdf", null, null],
  ["AVN-001", "Avon Technologies", "H1 2026 Results", "13-May-2026", "Latest Avon financials", "https://www.investegate.co.uk/announcement/rns/avon-technologies-plc--avon/interim-results/9565083", null, null],
  ["MKT-001", "Market data", "Closing-price references", "24-Sep-2026", "Screening market values", "https://stockanalysis.com/", null, null],
  ["FX-001", "Market data", "GBP/USD historical rate", "24-Sep-2026", "Avon translation", "https://www.exchangerates.org.uk/historical/GBP/24_09_2026", null, null],
  ["Internal", "Project", "Research files 01-08", "27-Sep-2026", "Scoring, valuation and purchase-accounting rationale", "See repository research folder", null, null],
];
header(sources, "C8:J8");
sources.getRange("C9:J16").format.wrapText = true;

section(sources, "C20:H20", "How the model works");
sources.getRange("C21:H27").values = [
  ["Step", "What happens", null, null, null, "Where to review"],
  [1, "The screening ranks the three targets using the agreed six categories.", null, null, null, "Screening"],
  [2, "Editable offer, funding, financing and synergy inputs are entered once.", null, null, null, "Assumptions"],
  [3, "The offer premium produces equity value and enterprise value.", null, null, null, "Deal"],
  [4, "A five-year forecast produces a standalone DCF and synergy-supported price ceiling.", null, null, null, "Valuation"],
  [5, "A simple purchase-price allocation estimates new intangibles, goodwill and amortisation.", null, null, null, "Purchase Accounting"],
  [6, "Earnings and leverage show the effect of financing, synergies and purchase accounting.", null, null, null, "Earnings / Leverage"],
];
header(sources, "C21:H21");
sources.getRange("C22:H27").format.wrapText = false;

section(sources, "C30:H30", "Current limitations");
sources.getRange("C31:H35").values = [
  ["The Cohort forecast is deliberately simple and is not a full three-statement model.", null, null, null, null, null],
  ["The model includes a DCF but not yet a full comparable-company or precedent-transactions analysis.", null, null, null, null, null],
  ["Purchase accounting is illustrative; it does not replace a detailed asset-by-asset valuation or tax review.", null, null, null, null, null],
  ["Net debt / EBIT is a leverage proxy, not the final rating-agency measure.", null, null, null, null, null],
  ["Market prices and exchange rates are point-in-time references and should be refreshed before use.", null, null, null, null, null],
];
sources.getRange("C31:H35").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };
sources.getRange("C31:H35").format.wrapText = false;

// Widths, row heights, freezes and tabs
for (const sheet of [summary, assumptions, valuation, purchaseAccounting, screening, deal, earnings, leverage, sources]) {
  sheet.getRange("A:B").format.columnWidth = 3;
  sheet.getRange("C:C").format.columnWidth = 30;
  sheet.getRange("D:D").format.columnWidth = 18;
  sheet.getRange("E:E").format.columnWidth = 14;
  sheet.getRange("F:F").format.columnWidth = 24;
  sheet.getRange("G:J").format.columnWidth = 18;
  sheet.getRange("2:4").format.rowHeight = 22;
}
summary.getRange("F:H").format.columnWidth = 25;
summary.getRange("C:C").format.columnWidth = 44;
summary.getRange("D:D").format.columnWidth = 38;
summary.getRange("E:E").format.columnWidth = 32;
earnings.getRange("C:C").format.columnWidth = 48;
earnings.getRange("D:D").format.columnWidth = 20;
assumptions.getRange("F:F").format.columnWidth = 42;
assumptions.getRange("G:G").format.columnWidth = 18;
valuation.getRange("C:C").format.columnWidth = 34;
valuation.getRange("D:E").format.columnWidth = 16;
valuation.getRange("F:F").format.columnWidth = 34;
valuation.getRange("G:G").format.columnWidth = 24;
valuation.getRange("H:I").format.columnWidth = 16;
purchaseAccounting.getRange("C:C").format.columnWidth = 42;
purchaseAccounting.getRange("D:D").format.columnWidth = 20;
purchaseAccounting.getRange("F:G").format.columnWidth = 18;
purchaseAccounting.getRange("H:I").format.columnWidth = 22;
purchaseAccounting.getRange("J:J").format.columnWidth = 25;
purchaseAccounting.getRange("K:K").format.columnWidth = 28;
screening.getRange("C:C").format.columnWidth = 34;
screening.getRange("H:H").format.columnWidth = 26;
screening.getRange("J:J").format.columnWidth = 36;
deal.getRange("F:F").format.columnWidth = 38;
sources.getRange("C:C").format.columnWidth = 12;
sources.getRange("D:D").format.columnWidth = 24;
sources.getRange("E:E").format.columnWidth = 32;
sources.getRange("F:F").format.columnWidth = 16;
sources.getRange("G:G").format.columnWidth = 30;
sources.getRange("H:H").format.columnWidth = 70;
summary.getRange("20:24").format.rowHeight = 34;
screening.getRange("26:29").format.rowHeight = 22;
sources.getRange("9:16").format.rowHeight = 46;
sources.getRange("22:27").format.rowHeight = 24;
sources.getRange("31:35").format.rowHeight = 22;
assumptions.freezePanes.freezeRows(7);
valuation.freezePanes.freezeRows(8);
purchaseAccounting.freezePanes.freezeRows(8);
screening.freezePanes.freezeRows(7);
sources.freezePanes.freezeRows(8);
summary.tabColor = navy;
assumptions.tabColor = "#5B9BD5";
valuation.tabColor = "#2F75B5";
purchaseAccounting.tabColor = "#70AD47";
screening.tabColor = "#4472C4";
sources.tabColor = "#A5A5A5";

wb.recalculate();

const summaryCheck = await wb.inspect({ kind: "table", range: "Summary!C7:H24", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 });
console.log(summaryCheck.ndjson);
const dealCheck = await wb.inspect({ kind: "table", range: "Deal!C8:G24", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 });
console.log(dealCheck.ndjson);
const valuationCheck = await wb.inspect({ kind: "table", range: "Valuation!C8:I37", include: "values,formulas", tableMaxRows: 40, tableMaxCols: 8 });
console.log(valuationCheck.ndjson);
const purchaseAccountingCheck = await wb.inspect({ kind: "table", range: "Purchase Accounting!C8:K27", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 10 });
console.log(purchaseAccountingCheck.ndjson);

// Input-response test in memory. Restore the base case before final verification and export.
assumptions.getRange("D21:D22").values = [[0.40], [0.80]];
wb.recalculate();
const responseTest = await wb.inspect({ kind: "table", range: "Summary!C8:D15", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 3 });
console.log(responseTest.ndjson);
assumptions.getRange("D21:D22").values = [[0.30], [0.60]];
wb.recalculate();

// Valuation input-response test. Restore the base case before final verification and export.
assumptions.getRange("G29").values = [[0.10]];
wb.recalculate();
const valuationResponseTest = await wb.inspect({ kind: "table", range: "Valuation!C30:G34", include: "values,formulas", tableMaxRows: 8, tableMaxCols: 6 });
console.log(valuationResponseTest.ndjson);
assumptions.getRange("G29").values = [[0.09]];
wb.recalculate();

// Purchase-accounting input-response test. Restore the base case before final verification and export.
assumptions.getRange("D39").values = [[0.45]];
wb.recalculate();
const purchaseAccountingResponseTest = await wb.inspect({ kind: "table", range: "Purchase Accounting!C11:D18", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 3 });
console.log(purchaseAccountingResponseTest.ndjson);
assumptions.getRange("D39").values = [[0.35]];
wb.recalculate();

const errorCheck = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log(errorCheck.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const sheetName of ["Summary", "Assumptions", "Valuation", "Purchase Accounting", "Screening", "Deal", "Earnings", "Leverage", "Sources & Guide"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "_").replaceAll("&", "and")}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
