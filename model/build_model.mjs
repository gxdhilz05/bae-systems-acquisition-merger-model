import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/workspace/scratch/d8279fcea226/bae-acquisition-merger-model/outputs/d8279fcea226";
const outputPath = `${outputDir}/bae_cohort_acquisition_model_v1.xlsx`;

const wb = Workbook.create();
const summary = wb.worksheets.add("Summary");
const assumptions = wb.worksheets.add("Assumptions");
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
  sheet.getRange("A1:K40").format.font = { name: fontFamily, size: 10, color: dark };
  sheet.getRange("A1:K40").format.verticalAlignment = "center";
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

for (const sheet of [summary, assumptions, screening, deal, earnings, leverage, sources]) baseSheet(sheet);

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
assumptions.getRange("F8:G16").values = [
  ["Input", "Value"],
  ["Cohort share price (p)", 1140.0],
  ["Cohort diluted shares (m)", 46.03],
  ["Cohort adjusted operating profit", 36.3],
  ["Cohort adjusted EPS (p)", 61.9],
  ["Cohort net debt / (funds)", -2.2],
  ["Cohort order book", 618.8],
  ["GBP per USD", 0.7565],
  ["Market data date", new Date("2026-09-24T00:00:00Z")],
];
header(assumptions, "F8:G8");
inputStyle(assumptions.getRange("G9:G16"));
assumptions.getRange("G9:G15").format.numberFormat = "#,##0.0;(#,##0.0);-";
assumptions.getRange("G15").format.numberFormat = "0.0000";
assumptions.getRange("G16").format.numberFormat = "dd-mmm-yy";

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
assumptions.getRange("C32:H34").values = [["How to use this tab", null, null, null, null, null], ["Change only the yellow blue-text cells. The offer, earnings and leverage schedules update automatically.", null, null, null, null, null], ["Base-case inputs are illustrative, not management guidance or a formal offer recommendation.", null, null, null, null, null]];
assumptions.getRange("C32:H32").format.font = { name: fontFamily, size: 10, bold: true, color: navy };
assumptions.getRange("C33:H34").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };

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
title(earnings, "Illustrative earnings impact", "Run-rate analysis excludes purchase-accounting amortisation and is not a full merger model.");
section(earnings, "C7:D7", "Run-rate earnings build");
earnings.getRange("C8:D22").values = [
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
crossLink(earnings.getRange("D9:D10"));
crossLink(earnings.getRange("D12:D15"));
earnings.getRange("D9:D10").format.numberFormat = "#,##0.0";
earnings.getRange("D11:D18").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
earnings.getRange("D19").format.numberFormat = "0.0p";
earnings.getRange("D20:D22").format.numberFormat = "0.0%";
totalStyle(earnings.getRange("C19:D20"));
earnings.getRange("F9:H14").values = [
  ["Interpretation", null, null],
  ["Positive accretion is not proof of value creation.", null, null],
  ["A premium can make EPS accretive while still earning a weak return.", null, null],
  ["This schedule excludes purchase-accounting amortisation.", null, null],
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
summary.getRange("C8:D15").values = [["Preferred target", "Cohort"], ["Cohort screening score", null], ["Illustrative offer price (p)", null], ["Illustrative equity purchase price", null], ["Illustrative enterprise value", null], ["Run-rate EPS accretion / (dilution)", null], ["Year 1 EPS impact incl. integration costs", null], ["Pro forma net debt / EBIT", null]];
summary.getRange("D9:D15").formulas = [["='Screening'!F14"],["='Deal'!D11"],["='Deal'!D13"],["='Deal'!D15"],["='Earnings'!D20"],["='Earnings'!D22"],["='Leverage'!D19"]];
summary.getRange("D9").format.numberFormat = "0";
summary.getRange("D10").format.numberFormat = "#,##0.0p";
summary.getRange("D11:D12").format.numberFormat = "#,##0.0;[Red](#,##0.0);-";
summary.getRange("D13:D14").format.numberFormat = "0.0%";
summary.getRange("D15").format.numberFormat = "0.00x";
summary.getRange("D8:D15").format.fill = paleBlue;
summary.getRange("C8:C15").format.font = { name: fontFamily, size: 10, bold: true, color: dark };
summary.getRange("F8:H15").values = [
  ["Current conclusion", null, null],
  ["Cohort provides the clearest link between specialist products and BAE platforms.", null, null],
  ["The proposed ownership model should preserve subsidiary autonomy.", null, null],
  ["The base case includes cost synergies but no revenue synergies.", null, null],
  ["EPS accretion alone does not establish value creation.", null, null],
  ["Purchase-accounting amortisation is not yet modelled.", null, null],
  ["Cash conversion, customer neutrality and regulatory approvals require diligence.", null, null],
  ["Next step: forecast Cohort and add valuation and purchase accounting.", null, null],
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
  ["Cash conversion", "FY2026 receivables reduced operating cash generation.", "Target forecast not yet built."],
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
  ["Internal", "Project", "Research files 01-06", "25-Sep-2026", "Scoring rationale and target selection", "See repository research folder", null, null],
];
header(sources, "C8:J8");
sources.getRange("C9:J16").format.wrapText = true;

section(sources, "C20:H20", "How the model works");
sources.getRange("C21:H27").values = [
  ["Step", "What happens", null, null, null, "Where to review"],
  [1, "The screening ranks the three targets using the agreed six categories.", null, null, null, "Screening"],
  [2, "Editable offer, funding, financing and synergy inputs are entered once.", null, null, null, "Assumptions"],
  [3, "The offer premium produces equity value and enterprise value.", null, null, null, "Deal"],
  [4, "Target earnings, synergies and financing costs produce an illustrative EPS impact.", null, null, null, "Earnings"],
  [5, "Transaction funding and target debt produce a pro forma leverage proxy.", null, null, null, "Leverage"],
  [6, "The Summary presents the answer and the conditions that could change it.", null, null, null, "Summary"],
];
header(sources, "C21:H21");
sources.getRange("C22:H27").format.wrapText = false;

section(sources, "C30:H30", "Current limitations");
sources.getRange("C31:H35").values = [
  ["A full Cohort forecast has not yet been built.", null, null, null, null, null],
  ["The model does not yet include a DCF, comparable-company valuation or premium analysis.", null, null, null, null, null],
  ["Purchase accounting, acquired-intangible amortisation and detailed integration phasing are not yet included.", null, null, null, null, null],
  ["Net debt / EBIT is a leverage proxy, not the final rating-agency measure.", null, null, null, null, null],
  ["Market prices and exchange rates are point-in-time references and should be refreshed before use.", null, null, null, null, null],
];
sources.getRange("C31:H35").format.font = { name: fontFamily, size: 10, italic: true, color: "#666666" };
sources.getRange("C31:H35").format.wrapText = false;

// Widths, row heights, freezes and tabs
for (const sheet of [summary, assumptions, screening, deal, earnings, leverage, sources]) {
  sheet.getRange("A:B").format.columnWidth = 3;
  sheet.getRange("C:C").format.columnWidth = 30;
  sheet.getRange("D:D").format.columnWidth = 18;
  sheet.getRange("E:E").format.columnWidth = 14;
  sheet.getRange("F:F").format.columnWidth = 24;
  sheet.getRange("G:J").format.columnWidth = 18;
  sheet.getRange("2:4").format.rowHeight = 22;
}
summary.getRange("F:H").format.columnWidth = 25;
summary.getRange("D:D").format.columnWidth = 38;
summary.getRange("E:E").format.columnWidth = 32;
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
screening.freezePanes.freezeRows(7);
sources.freezePanes.freezeRows(8);
summary.tabColor = navy;
assumptions.tabColor = "#5B9BD5";
screening.tabColor = "#4472C4";
sources.tabColor = "#A5A5A5";

wb.recalculate();

const summaryCheck = await wb.inspect({ kind: "table", range: "Summary!C7:H24", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 });
console.log(summaryCheck.ndjson);
const dealCheck = await wb.inspect({ kind: "table", range: "Deal!C8:G24", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 });
console.log(dealCheck.ndjson);

// Input-response test in memory. Restore the base case before final verification and export.
assumptions.getRange("D21:D22").values = [[0.40], [0.80]];
wb.recalculate();
const responseTest = await wb.inspect({ kind: "table", range: "Summary!C8:D15", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 3 });
console.log(responseTest.ndjson);
assumptions.getRange("D21:D22").values = [[0.30], [0.60]];
wb.recalculate();

const errorCheck = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log(errorCheck.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const sheetName of ["Summary", "Assumptions", "Screening", "Deal", "Earnings", "Leverage", "Sources & Guide"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "_").replaceAll("&", "and")}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
