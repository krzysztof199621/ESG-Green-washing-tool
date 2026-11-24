'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const pricing = require('./pricing.js');

const { calculateOneOff, calculateSubscription, formatQuoteText, OFFER_DATA } = pricing;

test('base prices, limits and timelines for all 5 one-off packages (manual mode)', () => {
  const expected = [
    { id: 'przeglad', price: 6900, evidence: 40, days: 5, workshop: 0 },
    { id: 'kampania', price: 12500, evidence: 40, days: 7, workshop: 0 },
    { id: 'compact', price: 16500, evidence: 80, days: 7, workshop: 60 },
    { id: 'standard', price: 23000, evidence: 80, days: 10, workshop: 90 },
    { id: 'enterprise', price: 39000, evidence: 80, days: 15, workshop: 180 }
  ];

  for (const exp of expected) {
    const res = calculateOneOff({ mode: 'manual', packageId: exp.id });
    assert.equal(res.selectedPackage.id, exp.id);
    assert.equal(res.totals.netto, exp.price);
    assert.equal(res.delivery.evidenceLimitPages, exp.evidence);
    assert.equal(res.delivery.baseWorkingDays, exp.days);
    assert.equal(res.delivery.workshopMinutes, exp.workshop);
  }
});

test('auto-selection for materials chooses cheapest valid package', () => {
  // 6 creatives -> Przeglad (6 900)
  let res = calculateOneOff({ mode: 'scope', scopeType: 'material', creatives: 6, packagings: 0 });
  assert.equal(res.selectedPackage.id, 'przeglad');
  assert.equal(res.totals.netto, 6900);

  // 10 creatives -> Przeglad (6 900) + 4*700 = 9 700 (cheaper than Kampania 12 500)
  res = calculateOneOff({ mode: 'scope', scopeType: 'material', creatives: 10, packagings: 0 });
  assert.equal(res.selectedPackage.id, 'przeglad');
  assert.equal(res.totals.netto, 9700);

  // 15 creatives -> Kampania (12 500, cheaper than Przeglad 6900 + 9*700 = 13 200)
  res = calculateOneOff({ mode: 'scope', scopeType: 'material', creatives: 15, packagings: 0 });
  assert.equal(res.selectedPackage.id, 'kampania');
  assert.equal(res.totals.netto, 12500);

  // 3 packagings -> Przeglad (6 900) + 2*2400 = 11 700
  res = calculateOneOff({ mode: 'scope', scopeType: 'material', creatives: 0, packagings: 3 });
  assert.equal(res.selectedPackage.id, 'przeglad');
  assert.equal(res.totals.netto, 11700);

  // 4 packagings -> Kampania (12 500, cheaper than Przeglad 6900 + 3*2400 = 14 100)
  res = calculateOneOff({ mode: 'scope', scopeType: 'material', creatives: 0, packagings: 4 });
  assert.equal(res.selectedPackage.id, 'kampania');
  assert.equal(res.totals.netto, 12500);

  // Campaign page + 5 creatives -> Kampania (12 500 vs Przeglad 6900 + 5*700 = 10 400 -> Przeglad is actually cheaper! Wait: check page+6 creatives vs page+10)
  // For campaign page + 5 creatives: Przeglad = 6900 + 5*700 = 10400, Kampania = 12500.
  res = calculateOneOff({ mode: 'scope', scopeType: 'material', hasCampaignPage: true, creatives: 5 });
  assert.equal(res.selectedPackage.id, 'przeglad');
  assert.equal(res.totals.netto, 10400);

  // For campaign page + 10 creatives: Przeglad = 6900 + 10*700 = 13900, Kampania = 12500 + 5*700 = 16000 -> still Przeglad;
  // Wait: Kampania includes page + 5 creatives, so Kampania base is 12500 while Przeglad + 5*700 is 10400.
  // If user selects Kampania manually for page + 5 creatives, price is 12500 with 0 extras.
  const manualKampania = calculateOneOff({ mode: 'manual', packageId: 'kampania' });
  assert.equal(manualKampania.totals.netto, 12500);
});

test('auto-selection for websites covers Compact, Standard, Enterprise, extra sites and >30 subpages warning', () => {
  // 8 subpages -> Compact (16 500)
  let res = calculateOneOff({ mode: 'scope', scopeType: 'website', sites: 1, subpages: 8 });
  assert.equal(res.selectedPackage.id, 'compact');
  assert.equal(res.totals.netto, 16500);

  // 12 subpages -> Standard (23 000)
  res = calculateOneOff({ mode: 'scope', scopeType: 'website', sites: 1, subpages: 12 });
  assert.equal(res.selectedPackage.id, 'standard');
  assert.equal(res.totals.netto, 23000);

  // 25 subpages -> Enterprise (39 000)
  res = calculateOneOff({ mode: 'scope', scopeType: 'website', sites: 1, subpages: 25 });
  assert.equal(res.selectedPackage.id, 'enterprise');
  assert.equal(res.totals.netto, 39000);

  // 3 sites, 20 subpages -> Enterprise (39 000) + 1 extra site (14 000) = 53 000
  res = calculateOneOff({ mode: 'scope', scopeType: 'website', sites: 3, subpages: 20 });
  assert.equal(res.selectedPackage.id, 'enterprise');
  assert.equal(res.totals.netto, 53000);

  // 40 subpages on 1 site -> Enterprise + SUBPAGES_OVER_30 warning
  res = calculateOneOff({ mode: 'scope', scopeType: 'website', sites: 1, subpages: 40 });
  assert.equal(res.selectedPackage.id, 'enterprise');
  assert.ok(res.warnings.some((w) => w.code === 'SUBPAGES_OVER_30'));
});

test('percentage surcharges, flat addons, 50/50 payment split and VAT 23%', () => {
  // Standard (23 000) + English (+25% = 5 750) + Express (+40% = 9 200) + Dictionary (8 000) + 2h workshop (1 200) + 3h revision (1 500)
  const res = calculateOneOff({
    mode: 'scope',
    scopeType: 'website',
    sites: 1,
    subpages: 15,
    english: true,
    express: true,
    dictionary: true,
    extraWorkshopHours: 2,
    extraRevisionHours: 3,
    vatRate: 0.23
  });

  assert.equal(res.selectedPackage.id, 'standard');
  assert.equal(res.auditScopeNetto, 23000);
  // 23000 + 5750 + 9200 + 8000 + 1200 + 1500 = 48 650
  assert.equal(res.totals.netto, 48650);
  assert.equal(res.totals.vatAmount, 11189.5);
  assert.equal(res.totals.brutto, 59839.5);
  assert.equal(res.paymentSchedule.tranche1Netto, 24325);
  assert.equal(res.paymentSchedule.tranche2Netto, 24325);
  assert.equal(res.delivery.workingDaysNumeric, 5);
  assert.equal(res.delivery.workshopMinutes, 90 + 120);
  assert.ok(res.warnings.some((w) => w.code === 'EXPRESS_CONFIRMATION'));
});

test('evidence page limit warnings for 40 and 80 pages', () => {
  const matWarn = calculateOneOff({ mode: 'manual', packageId: 'kampania', evidencePages: 41 });
  assert.ok(matWarn.warnings.some((w) => w.code === 'EVIDENCE_LIMIT_EXCEEDED'));

  const matOk = calculateOneOff({ mode: 'manual', packageId: 'kampania', evidencePages: 40 });
  assert.equal(matOk.warnings.some((w) => w.code === 'EVIDENCE_LIMIT_EXCEEDED'), false);

  const webWarn = calculateOneOff({ mode: 'manual', packageId: 'standard', evidencePages: 81 });
  assert.ok(webWarn.warnings.some((w) => w.code === 'EVIDENCE_LIMIT_EXCEEDED'));
});

test('subscription unit weights, rollover 30%, overage rates and auto plan selection', () => {
  // 2 posts (2) + 4 resizes (1) + 1 packaging (3) + 1 campaign (6) = 12 units
  const res = calculateSubscription({
    planId: 'auto',
    contractType: 'annual_monthly',
    brandsCount: 1,
    units: {
      postBanner: 2,
      resizeVariant: 4,
      packagingArticle: 1,
      campaignPresentation: 1
    }
  });

  assert.equal(res.totalUnits, 12);
  // Starter: 4900 + (12 - 8)*480 = 6820; Pro: 11900; Agency: 17900 -> Starter is cheapest
  assert.equal(res.selectedPlan.planId, 'starter');
  assert.equal(res.totals.monthlyNetto, 6820);
  assert.equal(res.selectedPlan.maxRolloverUnits, 2.4);

  // High volume: 5 campaigns (30) + 10 posts (10) = 40 units -> Starter: 4900 + 32*480 = 20260, Pro: 11900 + 20*450 = 20900, Agency: 17900
  const highVol = calculateSubscription({
    planId: 'auto',
    contractType: 'annual_monthly',
    brandsCount: 1,
    units: { postBanner: 10, campaignPresentation: 5 }
  });
  assert.equal(highVol.totalUnits, 40);
  assert.equal(highVol.selectedPlan.planId, 'agency');
  assert.equal(highVol.totals.monthlyNetto, 17900);
  assert.equal(highVol.selectedPlan.maxRolloverUnits, 12);
});

test('subscription contract variants: annual upfront (-8%), 3m pilot (+20% & credit), 3rd brand Pro and Agency urgent 6h', () => {
  // Pro with 3 brands (+3000 -> base 14 900) and annual_upfront (-8% -> 13 708/mies.)
  const proUpfront = calculateSubscription({
    planId: 'pro',
    contractType: 'annual_upfront',
    brandsCount: 3,
    units: { postBanner: 10 }
  });
  assert.equal(proUpfront.selectedPlan.extraBrandFee, 3000);
  assert.equal(proUpfront.selectedPlan.baseMonthlyRecurring, 14900);
  assert.equal(proUpfront.selectedPlan.adjustedMonthlyRecurring, 13708);
  assert.equal(proUpfront.selectedPlan.upfrontSavingsNetto, 14304);
  assert.equal(proUpfront.selectedPlan.rolloverUnits, 6); // 10 unused capped at 30% of 20 = 6

  // Agency with urgent 6h (+6000 -> base 23 900) and pilot_3m (+20% -> 28 680/mies., credit 3 * 23900 * 0.2 = 14 340)
  const agencyPilot = calculateSubscription({
    planId: 'agency',
    contractType: 'pilot_3m',
    brandsCount: 5,
    urgent6h: true,
    units: { postBanner: 20 }
  });
  assert.equal(agencyPilot.selectedPlan.urgent6hFee, 6000);
  assert.equal(agencyPilot.selectedPlan.adjustedMonthlyRecurring, 28680);
  assert.equal(agencyPilot.selectedPlan.pilotConversionCreditNetto, 14340);
  assert.equal(agencyPilot.totals.contractNetto, 86040);
});

test('subscription website discounts (annual vs pilot), Pro extra diff scans and out-of-pool items', () => {
  // Pro annual + Standard website (23 000 * 0.6 = 13 800) + 1 extra diff scan (5 500) + 1 out-of-pool Compact (16 500)
  const proAnnualWeb = calculateSubscription({
    planId: 'pro',
    contractType: 'annual_monthly',
    brandsCount: 2,
    units: { postBanner: 20 },
    websiteAuditPackage: 'standard',
    extraDiffScans: 1,
    outOfPoolCompact: 1
  });
  assert.equal(proAnnualWeb.websiteDiscountPct, 0.40);
  assert.equal(proAnnualWeb.totals.oneOffExtrasNetto, 13800 + 5500 + 16500);

  // Starter pilot + Compact website -> no 30% discount in pilot
  const starterPilotWeb = calculateSubscription({
    planId: 'starter',
    contractType: 'pilot_3m',
    brandsCount: 1,
    units: { postBanner: 4 },
    websiteAuditPackage: 'compact'
  });
  assert.equal(starterPilotWeb.websiteDiscountPct, 0);
  assert.equal(starterPilotWeb.totals.oneOffExtrasNetto, 16500);

  // Agency pilot + Compact website -> 25% discount applies
  const agencyPilotWeb = calculateSubscription({
    planId: 'agency',
    contractType: 'pilot_3m',
    brandsCount: 4,
    units: { postBanner: 10 },
    websiteAuditPackage: 'compact'
  });
  assert.equal(agencyPilotWeb.websiteDiscountPct, 0.25);
  assert.equal(agencyPilotWeb.totals.oneOffExtrasNetto, 12375);
});

test('formatQuoteText renders Polish summary with client metadata', () => {
  const res = calculateOneOff({ mode: 'manual', packageId: 'przeglad' });
  const text = formatQuoteText(res, {
    clientName: 'Firma Testowa Sp. z o.o.',
    projectName: 'Kampania Wiosna',
    contactPerson: 'Jan Nowak'
  });
  assert.match(text, /Firma Testowa Sp\. z o\.o\./);
  assert.match(text, /Audyt materiału przed publikacją/);
  assert.match(text, /6\s?900 zł/);
  assert.match(text, /Agnieszka Orłowska/);
  assert.equal(OFFER_DATA.EXCLUSIONS.length, 4);
});
