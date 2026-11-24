(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.ESGPricing = factory();
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var OFFER_DATA = {
    meta: {
      company: 'ESG Institute sp. z o.o.',
      address: 'Rondo ONZ 1 p.10, 00-124 Warszawa',
      website: 'esginstitute.eu',
      contactName: 'Agnieszka Orłowska',
      contactEmail: 'orlowska@esginstitute.eu',
      contactPhone: '+48 722 230 101',
      offerDate: '16 listopada 2026 r.',
      validUntil: '15 stycznia 2027 r.',
      defaultVatRate: 0.23,
      revisionDaysLimit: 5,
      includedRevisionRounds: 1,
      overagePaymentDays: 14,
      pilotConversionWindowDays: 14,
      empcoDatePL: '27 września 2027 r.',
      empcoDateDE: '27 września 2026 r.'
    },

    ONE_OFF_PACKAGES: [
      {
        id: 'przeglad',
        name: 'Audyt materiału przed publikacją',
        shortName: 'Przegląd przed publikacją',
        category: 'material',
        priceNetto: 6900,
        evidenceLimitPages: 40,
        workshopMinutes: 0,
        workingDays: 5,
        includedSummary: 'Jeden materiał: strona kampanii, do 6 haseł lub kreacji, albo jedno opakowanie.',
        limits: {
          campaignPages: 1,
          creativesStandalone: 6,
          creativesWithPage: 0,
          packagingsStandalone: 1
        }
      },
      {
        id: 'kampania',
        name: 'Audyt kampanii lub opakowania',
        shortName: 'Kampania / opakowanie',
        category: 'material',
        priceNetto: 12500,
        evidenceLimitPages: 40,
        workshopMinutes: 0,
        workingDays: 7,
        includedSummary: 'Do 15 kreacji albo do 4 opakowań, albo strona kampanii i 5 kreacji.',
        limits: {
          campaignPages: 1,
          creativesStandalone: 15,
          creativesWithPage: 5,
          packagingsStandalone: 4
        }
      },
      {
        id: 'compact',
        name: 'Audyt strony www Compact',
        shortName: 'Strona www Compact',
        category: 'website',
        priceNetto: 16500,
        evidenceLimitPages: 80,
        workshopMinutes: 60,
        workingDays: 7,
        includedSummary: 'Do 8 podstron (1 serwis). Warsztat z ekspertem 60 minut.',
        limits: {
          sites: 1,
          subpages: 8
        }
      },
      {
        id: 'standard',
        name: 'Audyt strony www Standard',
        shortName: 'Strona www Standard',
        category: 'website',
        priceNetto: 23000,
        evidenceLimitPages: 80,
        workshopMinutes: 90,
        workingDays: 10,
        includedSummary: 'Do 15 podstron (1 serwis). Warsztat z ekspertem 90 minut.',
        limits: {
          sites: 1,
          subpages: 15
        }
      },
      {
        id: 'enterprise',
        name: 'Audyt strony www Enterprise',
        shortName: 'Strona www Enterprise',
        category: 'website',
        priceNetto: 39000,
        evidenceLimitPages: 80,
        workshopMinutes: 180,
        workingDays: 15,
        includedSummary: 'Do 2 serwisów albo 30 podstron. Warsztat z ekspertem 3 godziny.',
        limits: {
          sites: 2,
          subpages: 30
        }
      }
    ],

    ONE_OFF_ADDONS: {
      extraCreative: 700,
      extraPackaging: 2400,
      extraSite: 14000,
      dictionary: 8000,
      englishRate: 0.25,
      expressRate: 0.40,
      extraWorkshopHour: 600,
      extraRevisionHour: 500
    },

    SUBSCRIPTION_PLANS: [
      {
        id: 'starter',
        name: 'Starter',
        monthlyNetto: 4900,
        audience: 'Jedna marka albo jedna spółka.',
        maxBrands: 1,
        includedBrands: 1,
        extraBrandFee: null,
        unitsPool: 8,
        rolloverPct: 0.30,
        overageRate: 480,
        responseDays: 3,
        websiteDiscountAnnual: 0.30,
        websiteDiscountPilot: 0,
        websiteDiscountLabel: 'Poza abonamentem. Przy umowie rocznej rabat 30% na jeden audyt z cennika.',
        includedDiffScans: 0,
        extraDiffScanPrice: null,
        urgent6hFee: null,
        urgent6hMaxMaterials: 0,
        perks: []
      },
      {
        id: 'pro',
        name: 'Pro',
        monthlyNetto: 11900,
        audience: 'Do dwóch marek. Trzecia marka: +3 000 zł miesięcznie.',
        maxBrands: 3,
        includedBrands: 2,
        extraBrandFee: 3000,
        unitsPool: 20,
        rolloverPct: 0.30,
        overageRate: 450,
        responseDays: 2,
        websiteDiscountAnnual: 0.40,
        websiteDiscountPilot: 0,
        websiteDiscountLabel: 'Pełen audyt strony www przy umowie rocznej z rabatem 40%. Po pełnym audycie 2 skany różnicowe w roku (do 15 podstron). Kolejny skan: 5 500 zł.',
        includedDiffScans: 2,
        extraDiffScanPrice: 5500,
        urgent6hFee: null,
        urgent6hMaxMaterials: 0,
        perks: ['2 skany różnicowe strony w roku po pełnym audycie (do 15 podstron)']
      },
      {
        id: 'agency',
        name: 'Agency',
        monthlyNetto: 17900,
        audience: 'Agencja reklamowa lub PR. W jednym miesiącu do 8 marek klientów.',
        maxBrands: 8,
        includedBrands: 8,
        extraBrandFee: null,
        unitsPool: 40,
        rolloverPct: 0.30,
        overageRate: 400,
        responseDays: 1,
        websiteDiscountAnnual: 0.25,
        websiteDiscountPilot: 0.25,
        websiteDiscountLabel: 'Audyty stron klientów z cennika projektowego z rabatem 25%.',
        includedDiffScans: 0,
        extraDiffScanPrice: null,
        urgent6hFee: 6000,
        urgent6hMaxMaterials: 6,
        perks: [
          'Raport z logo agencji',
          'Szkolenie zespołu na start (pół dnia)'
        ]
      }
    ],

    SUBSCRIPTION_CONTRACTS: {
      annual_monthly: {
        id: 'annual_monthly',
        name: 'Umowa roczna (płatność miesięczna)',
        months: 12,
        multiplier: 1.0,
        discountPct: 0,
        surchargePct: 0,
        description: 'Współpraca na 12 miesięcy, rozliczenie miesięczne płatne z góry.'
      },
      annual_upfront: {
        id: 'annual_upfront',
        name: 'Umowa roczna (płatność z góry, -8%)',
        months: 12,
        multiplier: 0.92,
        discountPct: 0.08,
        surchargePct: 0,
        description: 'Płatność jednorazowa za cały rok z góry z rabatem 8%.'
      },
      pilot_3m: {
        id: 'pilot_3m',
        name: 'Program pilotażowy (3 miesiące, +20%)',
        months: 3,
        multiplier: 1.20,
        discountPct: 0,
        surchargePct: 0.20,
        description: '3-miesięczny okres próbny (stawka +20%). Nadpłacona różnica zaliczana na 1. fakturę umowy 12-mies. przy decyzji w ciągu 14 dni.'
      }
    },

    UNIT_TYPES: [
      {
        id: 'postBanner',
        units: 1,
        label: 'Post, baner, newsletter lub hasła',
        description: 'Post, baner, newsletter do 2 000 znaków albo do dwóch haseł.'
      },
      {
        id: 'resizeVariant',
        units: 0.25,
        label: 'Kolejny wymiar tej samej kreacji',
        description: 'Kolejny wymiar tej samej kreacji, bez zmiany tekstu.'
      },
      {
        id: 'packagingArticle',
        units: 3,
        label: 'Etykieta / opakowanie, ulotka, artykuł, wideo',
        description: 'Etykieta lub opakowanie jednego produktu, ulotka, artykuł do 6 000 znaków, scenariusz nagrania do 60 sekund.'
      },
      {
        id: 'campaignPresentation',
        units: 6,
        label: 'Strona kampanii, spot, prezentacja / broszura',
        description: 'Strona kampanii z podlinkowanymi certyfikatami, scenariusz spotu, prezentacja lub broszura do 10 stron.'
      }
    ],

    OUT_OF_POOL: {
      compactPrice: 16500,
      standardPrice: 23000,
      description: 'Sekcja serwisu powyżej 3 podstron, katalog, ebook albo raport ponad 10 stron. Cena jak Compact (16 500 zł) albo Standard (23 000 zł).'
    },

    EXCLUSIONS: [
      'Atestacja raportów ESG oraz badanie sprawozdań finansowych.',
      'Odtwarzanie analizy cyklu życia (LCA) ani obliczanie śladu węglowego.',
      'Przygotowywanie pism do UOKiK oraz opinii na potrzeby postępowań spornych.',
      'Gwarancja braku zakwestionowania komunikacji przez organy regulacyjne.'
    ],

    RISK_STATUSES: [
      {
        status: 'Czerwony',
        action: 'Do usunięcia / zastąpienia',
        detail: 'Twierdzenie wprowadza w błąd w świetle UPNPR / UoZNK (od 27.09.2027 r. także czarna lista EmpCo UE 2024/825). Zdjęcie materiału; zdanie zastępcze wyłącznie przy pełnym pokryciu w dowodach.'
      },
      {
        status: 'Pomarańczowy',
        action: 'Do zawężenia',
        detail: 'Twierdzenie ma oparcie w dokumentach firmy, ale jest szersze niż dowód (np. certyfikat jednej spółki przypisany grupie). Wstawienie zdania ograniczonego do właściwego podmiotu, zakresu i okresu.'
      }
    ]
  };

  function round2(value) {
    return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  }

  function toInt(value, fallback) {
    var n = Number(value);
    if (!Number.isFinite(n) || n < 0) return fallback || 0;
    return Math.floor(n);
  }

  function toNum(value, fallback) {
    var n = Number(value);
    if (!Number.isFinite(n) || n < 0) return fallback || 0;
    return n;
  }

  function getPackageById(id) {
    for (var i = 0; i < OFFER_DATA.ONE_OFF_PACKAGES.length; i++) {
      if (OFFER_DATA.ONE_OFF_PACKAGES[i].id === id) {
        return OFFER_DATA.ONE_OFF_PACKAGES[i];
      }
    }
    return null;
  }

  function getPlanById(id) {
    for (var i = 0; i < OFFER_DATA.SUBSCRIPTION_PLANS.length; i++) {
      if (OFFER_DATA.SUBSCRIPTION_PLANS[i].id === id) {
        return OFFER_DATA.SUBSCRIPTION_PLANS[i];
      }
    }
    return null;
  }

  function formatPLN(amount) {
    var rounded = round2(amount);
    var hasFraction = Math.abs(rounded - Math.round(rounded)) > 0.001;
    return rounded.toLocaleString('pl-PL', {
      minimumFractionDigits: hasFraction ? 2 : 0,
      maximumFractionDigits: 2
    }) + ' zł';
  }

  function formatUnits(units) {
    var rounded = round2(units);
    return rounded.toLocaleString('pl-PL', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2
    });
  }

  function formatExpressDays(baseDays) {
    var half = baseDays / 2;
    if (Number.isInteger(half)) {
      var unit = half === 1 ? 'dzień roboczy' : (half < 5 ? 'dni robocze' : 'dni roboczych');
      return {
        numeric: half,
        label: half + ' ' + unit
      };
    }
    var low = Math.floor(half);
    var high = Math.ceil(half);
    return {
      numeric: half,
      label: low + '–' + high + ' dni robocze'
    };
  }

  function formatNormalDays(days) {
    return {
      numeric: days,
      label: days + ' ' + (days === 1 ? 'dzień roboczy' : (days < 5 ? 'dni robocze' : 'dni roboczych'))
    };
  }

  function formatWorkshop(minutes) {
    if (!minutes) return 'Brak w pakiecie';
    if (minutes % 60 === 0) {
      var h = minutes / 60;
      return h + ' ' + (h === 1 ? 'godzina' : (h < 5 ? 'godziny' : 'godzin'));
    }
    return minutes + ' minut';
  }

  function evaluateMaterialOption(pkgId, hasCampaignPage, creatives, packagings) {
    var pkg = getPackageById(pkgId);
    var addons = OFFER_DATA.ONE_OFF_ADDONS;
    var extraCreatives = 0;
    var extraPackagings = 0;
    var allocationNote = null;

    if (pkgId === 'przeglad') {
      if (hasCampaignPage) {
        extraCreatives = creatives;
        extraPackagings = packagings;
      } else if (creatives > 0 && packagings === 0) {
        extraCreatives = Math.max(0, creatives - 6);
      } else if (packagings > 0 && creatives === 0) {
        extraPackagings = Math.max(0, packagings - 1);
      } else if (creatives > 0 && packagings > 0) {
        var extraCreativesA = Math.max(0, creatives - 6);
        var extraPackagingsA = packagings;
        var costA = extraCreativesA * addons.extraCreative + extraPackagingsA * addons.extraPackaging;

        var extraCreativesB = creatives;
        var extraPackagingsB = Math.max(0, packagings - 1);
        var costB = extraCreativesB * addons.extraCreative + extraPackagingsB * addons.extraPackaging;

        if (costB <= costA) {
          extraCreatives = extraCreativesB;
          extraPackagings = extraPackagingsB;
          allocationNote = 'Pakiet bazowy Przegląd obejmuje 1 materiał (przypisano 1 opakowanie jako bazę, pozostałe elementy doliczono jako dopłaty).';
        } else {
          extraCreatives = extraCreativesA;
          extraPackagings = extraPackagingsA;
          allocationNote = 'Pakiet bazowy Przegląd obejmuje 1 materiał (przypisano do 6 kreacji jako bazę, opakowania doliczono jako dopłaty).';
        }
      }
    } else if (pkgId === 'kampania') {
      if (hasCampaignPage) {
        extraCreatives = Math.max(0, creatives - 5);
        extraPackagings = packagings;
      } else if (creatives > 0 && packagings === 0) {
        extraCreatives = Math.max(0, creatives - 15);
      } else if (packagings > 0 && creatives === 0) {
        extraPackagings = Math.max(0, packagings - 4);
      } else if (creatives > 0 && packagings > 0) {
        var kExtraCreativesA = Math.max(0, creatives - 15);
        var kExtraPackagingsA = packagings;
        var kCostA = kExtraCreativesA * addons.extraCreative + kExtraPackagingsA * addons.extraPackaging;

        var kExtraCreativesB = creatives;
        var kExtraPackagingsB = Math.max(0, packagings - 4);
        var kCostB = kExtraCreativesB * addons.extraCreative + kExtraPackagingsB * addons.extraPackaging;

        if (kCostB <= kCostA) {
          extraCreatives = kExtraCreativesB;
          extraPackagings = kExtraPackagingsB;
          allocationNote = 'Pakiet Kampania obejmuje bazowo do 4 opakowań albo do 15 kreacji (przyjęto bazę opakowań jako korzystniejszą cenowo).';
        } else {
          extraCreatives = kExtraCreativesA;
          extraPackagings = kExtraPackagingsA;
          allocationNote = 'Pakiet Kampania obejmuje bazowo do 15 kreacji albo do 4 opakowań (przyjęto bazę kreacji jako korzystniejszą cenowo).';
        }
      }
    }

    var extrasCost = extraCreatives * addons.extraCreative + extraPackagings * addons.extraPackaging;
    return {
      packageId: pkg.id,
      packageName: pkg.name,
      shortName: pkg.shortName,
      category: pkg.category,
      eligible: true,
      basePriceNetto: pkg.priceNetto,
      discountedBaseNetto: pkg.priceNetto,
      extraCreatives: extraCreatives,
      extraPackagings: extraPackagings,
      extraSites: 0,
      scopeNetto: pkg.priceNetto + extrasCost,
      workingDays: pkg.workingDays,
      evidenceLimitPages: pkg.evidenceLimitPages,
      workshopMinutes: pkg.workshopMinutes,
      allocationNote: allocationNote
    };
  }

  function evaluateWebsiteOption(pkgId, sites, subpages, discountPct) {
    var pkg = getPackageById(pkgId);
    var addons = OFFER_DATA.ONE_OFF_ADDONS;
    var eligible = true;
    var extraSites = 0;

    if (pkgId === 'compact') {
      eligible = subpages <= 8;
      extraSites = Math.max(0, sites - 1);
    } else if (pkgId === 'standard') {
      eligible = subpages <= 15;
      extraSites = Math.max(0, sites - 1);
    } else if (pkgId === 'enterprise') {
      eligible = true;
      extraSites = Math.max(0, sites - 2);
    }

    var discountedBaseNetto = round2(pkg.priceNetto * (1 - discountPct));
    var extrasCost = extraSites * addons.extraSite;

    return {
      packageId: pkg.id,
      packageName: pkg.name,
      shortName: pkg.shortName,
      category: pkg.category,
      eligible: eligible,
      basePriceNetto: pkg.priceNetto,
      discountedBaseNetto: discountedBaseNetto,
      extraCreatives: 0,
      extraPackagings: 0,
      extraSites: extraSites,
      scopeNetto: round2(discountedBaseNetto + extrasCost),
      workingDays: pkg.workingDays,
      evidenceLimitPages: pkg.evidenceLimitPages,
      workshopMinutes: pkg.workshopMinutes,
      allocationNote: null
    };
  }

  function calculateOneOff(rawInput) {
    var input = rawInput || {};
    var addons = OFFER_DATA.ONE_OFF_ADDONS;
    var mode = input.mode === 'manual' ? 'manual' : 'scope';
    var scopeType = input.scopeType === 'website' ? 'website' : 'material';
    var vatRate = input.vatRate !== undefined ? toNum(input.vatRate, OFFER_DATA.meta.defaultVatRate) : OFFER_DATA.meta.defaultVatRate;
    var requestedDiscount = toNum(input.subscriptionDiscountPct, 0);
    if (requestedDiscount > 1) requestedDiscount = requestedDiscount / 100;

    var evidencePages = toInt(input.evidencePages, 0);
    var dictionary = Boolean(input.dictionary);
    var english = Boolean(input.english);
    var express = Boolean(input.express);
    var extraWorkshopHours = toNum(input.extraWorkshopHours, 0);
    var extraRevisionHours = toNum(input.extraRevisionHours, 0);

    var warnings = [];
    var notes = [];
    var alternatives = [];
    var selectedEval = null;

    if (mode === 'scope') {
      if (scopeType === 'material') {
        var hasCampaignPage = Boolean(input.hasCampaignPage);
        var creatives = toInt(input.creatives, 0);
        var packagings = toInt(input.packagings, 0);

        var optPrzeglad = evaluateMaterialOption('przeglad', hasCampaignPage, creatives, packagings);
        var optKampania = evaluateMaterialOption('kampania', hasCampaignPage, creatives, packagings);
        alternatives = [optPrzeglad, optKampania];

        if (optKampania.scopeNetto <= optPrzeglad.scopeNetto) {
          selectedEval = optKampania;
        } else {
          selectedEval = optPrzeglad;
        }
        if (selectedEval.allocationNote) {
          notes.push(selectedEval.allocationNote);
        }
      } else {
        var sites = Math.max(1, toInt(input.sites, 1));
        var subpages = Math.max(1, toInt(input.subpages, 1));

        var optCompact = evaluateWebsiteOption('compact', sites, subpages, requestedDiscount);
        var optStandard = evaluateWebsiteOption('standard', sites, subpages, requestedDiscount);
        var optEnterprise = evaluateWebsiteOption('enterprise', sites, subpages, requestedDiscount);

        var allWebOptions = [optCompact, optStandard, optEnterprise];
        alternatives = allWebOptions.filter(function (opt) { return opt.eligible; });

        selectedEval = alternatives[0];
        for (var w = 1; w < alternatives.length; w++) {
          if (alternatives[w].scopeNetto < selectedEval.scopeNetto) {
            selectedEval = alternatives[w];
          }
        }

        if ((subpages > 30 && sites <= 2) || (subpages > 30 && subpages > sites * 15)) {
          warnings.push({
            severity: 'warning',
            code: 'SUBPAGES_OVER_30',
            message: 'Zakres przekracza 30 podstron (pakiet Enterprise obejmuje do 2 serwisów albo 30 podstron). Wymagana indywidualna wycena lub dokupienie kolejnego serwisu przed rozpoczęciem prac.'
          });
        }
      }
    } else {
      var manualPkgId = input.packageId || 'przeglad';
      var manualPkg = getPackageById(manualPkgId) || getPackageById('przeglad');
      var manualExtraCreatives = toInt(input.extraCreatives, 0);
      var manualExtraPackagings = toInt(input.extraPackagings, 0);
      var manualExtraSites = toInt(input.extraSites, 0);
      var effectiveManualDiscount = manualPkg.category === 'website' ? requestedDiscount : 0;
      var manualBaseDiscounted = round2(manualPkg.priceNetto * (1 - effectiveManualDiscount));
      var manualExtrasCost =
        manualExtraCreatives * addons.extraCreative +
        manualExtraPackagings * addons.extraPackaging +
        manualExtraSites * addons.extraSite;

      selectedEval = {
        packageId: manualPkg.id,
        packageName: manualPkg.name,
        shortName: manualPkg.shortName,
        category: manualPkg.category,
        eligible: true,
        basePriceNetto: manualPkg.priceNetto,
        discountedBaseNetto: manualBaseDiscounted,
        extraCreatives: manualExtraCreatives,
        extraPackagings: manualExtraPackagings,
        extraSites: manualExtraSites,
        scopeNetto: round2(manualBaseDiscounted + manualExtrasCost),
        workingDays: manualPkg.workingDays,
        evidenceLimitPages: manualPkg.evidenceLimitPages,
        workshopMinutes: manualPkg.workshopMinutes,
        allocationNote: null
      };
      alternatives = [selectedEval];
    }

    var pkg = getPackageById(selectedEval.packageId);
    var appliedDiscountPct = selectedEval.category === 'website' ? requestedDiscount : 0;
    if (requestedDiscount > 0 && selectedEval.category !== 'website') {
      notes.push('Rabat abonamentowy na audyt strony WWW nie ma zastosowania do pakietów materiałów/kampanii — wycena bez rabatu.');
    }

    var lineItems = [];
    lineItems.push({
      id: 'base_package',
      label: pkg.name,
      detail: pkg.includedSummary,
      quantity: 1,
      unitPriceNetto: pkg.priceNetto,
      totalNetto: pkg.priceNetto
    });

    if (appliedDiscountPct > 0) {
      var discountAmount = round2(pkg.priceNetto - selectedEval.discountedBaseNetto);
      lineItems.push({
        id: 'subscription_discount',
        label: 'Rabat abonamentowy na audyt strony WWW (-' + Math.round(appliedDiscountPct * 100) + '%)',
        detail: 'Dotyczy wyłącznie pakietów audytu strony WWW zgodnie z warunkami planu abonamentowego.',
        quantity: 1,
        unitPriceNetto: -discountAmount,
        totalNetto: -discountAmount
      });
    }

    if (selectedEval.extraCreatives > 0) {
      lineItems.push({
        id: 'extra_creatives',
        label: 'Kolejne kreacje ponad limit pakietu',
        detail: selectedEval.extraCreatives + ' szt. × ' + formatPLN(addons.extraCreative),
        quantity: selectedEval.extraCreatives,
        unitPriceNetto: addons.extraCreative,
        totalNetto: round2(selectedEval.extraCreatives * addons.extraCreative)
      });
    }

    if (selectedEval.extraPackagings > 0) {
      lineItems.push({
        id: 'extra_packagings',
        label: 'Kolejne opakowania ponad limit pakietu',
        detail: selectedEval.extraPackagings + ' szt. × ' + formatPLN(addons.extraPackaging),
        quantity: selectedEval.extraPackagings,
        unitPriceNetto: addons.extraPackaging,
        totalNetto: round2(selectedEval.extraPackagings * addons.extraPackaging)
      });
    }

    if (selectedEval.extraSites > 0) {
      lineItems.push({
        id: 'extra_sites',
        label: 'Kolejne serwisy WWW ponad limit pakietu',
        detail: selectedEval.extraSites + ' szt. × ' + formatPLN(addons.extraSite),
        quantity: selectedEval.extraSites,
        unitPriceNetto: addons.extraSite,
        totalNetto: round2(selectedEval.extraSites * addons.extraSite)
      });
    }

    var auditScopeNetto = selectedEval.scopeNetto;
    var englishNetto = english ? round2(auditScopeNetto * addons.englishRate) : 0;
    var expressNetto = express ? round2(auditScopeNetto * addons.expressRate) : 0;
    var dictionaryNetto = dictionary ? addons.dictionary : 0;
    var extraWorkshopNetto = extraWorkshopHours > 0 ? round2(extraWorkshopHours * addons.extraWorkshopHour) : 0;
    var extraRevisionNetto = extraRevisionHours > 0 ? round2(extraRevisionHours * addons.extraRevisionHour) : 0;

    if (english) {
      lineItems.push({
        id: 'english',
        label: 'Werdykt i propozycje zdań po angielsku (+25%)',
        detail: '25% od wartości zakresu audytu (' + formatPLN(auditScopeNetto) + ')',
        quantity: 1,
        unitPriceNetto: englishNetto,
        totalNetto: englishNetto
      });
    }

    if (express) {
      lineItems.push({
        id: 'express',
        label: 'Tryb Ekspres: połowa czasu z oferty (+40%)',
        detail: '40% od wartości zakresu audytu (' + formatPLN(auditScopeNetto) + '), po potwierdzeniu terminu',
        quantity: 1,
        unitPriceNetto: expressNetto,
        totalNetto: expressNetto
      });
      warnings.push({
        severity: 'info',
        code: 'EXPRESS_CONFIRMATION',
        message: 'Tryb Ekspres (+40%) skraca realizację o połowę, ale wymaga indywidualnego potwierdzenia dostępności terminu.'
      });
    }

    if (dictionary) {
      lineItems.push({
        id: 'dictionary',
        label: 'Słownik dozwolonych twierdzeń komunikacyjnych (do 40 haseł)',
        detail: 'Powstaje z dowodów zebranych w audycie (Compact lub Standard).',
        quantity: 1,
        unitPriceNetto: addons.dictionary,
        totalNetto: dictionaryNetto
      });
      if (pkg.id !== 'compact' && pkg.id !== 'standard') {
        warnings.push({
          severity: 'warning',
          code: 'DICTIONARY_PACKAGE_MISMATCH',
          message: 'Wg cennika Słownik dozwolonych twierdzeń (8 000 zł, do 40 haseł) dedykowany jest po pakietach Compact lub Standard — dla wybranego pakietu może wymagać potwierdzenia.'
        });
      }
    }

    if (extraWorkshopHours > 0) {
      lineItems.push({
        id: 'extra_workshop',
        label: 'Dodatkowe godziny warsztatu z ekspertem',
        detail: formatUnits(extraWorkshopHours) + ' godz. × ' + formatPLN(addons.extraWorkshopHour),
        quantity: extraWorkshopHours,
        unitPriceNetto: addons.extraWorkshopHour,
        totalNetto: extraWorkshopNetto
      });
    }

    if (extraRevisionHours > 0) {
      lineItems.push({
        id: 'extra_revision',
        label: 'Dodatkowa runda konsultacji lub poprawek',
        detail: formatUnits(extraRevisionHours) + ' godz. × ' + formatPLN(addons.extraRevisionHour) + ' (1 runda uwag w ciągu 5 dni rob. jest w cenie)',
        quantity: extraRevisionHours,
        unitPriceNetto: addons.extraRevisionHour,
        totalNetto: extraRevisionNetto
      });
    }

    if (evidencePages > pkg.evidenceLimitPages) {
      warnings.push({
        severity: 'warning',
        code: 'EVIDENCE_LIMIT_EXCEEDED',
        message: 'Dokumentacja dowodowa (' + evidencePages + ' str.) przekracza limit pakietu (' + pkg.evidenceLimitPages + ' str.). Wymagana indywidualna wycena przed rozpoczęciem prac.'
      });
    }

    var totalNetto = round2(
      auditScopeNetto +
      englishNetto +
      expressNetto +
      dictionaryNetto +
      extraWorkshopNetto +
      extraRevisionNetto
    );
    var vatAmount = round2(totalNetto * vatRate);
    var totalBrutto = round2(totalNetto + vatAmount);

    var tranche1Netto = round2(totalNetto * 0.5);
    var tranche2Netto = round2(totalNetto - tranche1Netto);
    var tranche1Brutto = round2(totalBrutto * 0.5);
    var tranche2Brutto = round2(totalBrutto - tranche1Brutto);

    var timeline = express ? formatExpressDays(pkg.workingDays) : formatNormalDays(pkg.workingDays);
    var totalWorkshopMinutes = pkg.workshopMinutes + Math.round(extraWorkshopHours * 60);

    return {
      kind: 'one_off',
      mode: mode,
      scopeType: scopeType,
      selectedPackage: pkg,
      appliedDiscountPct: appliedDiscountPct,
      auditScopeNetto: auditScopeNetto,
      lineItems: lineItems,
      alternatives: alternatives,
      totals: {
        netto: totalNetto,
        vatRate: vatRate,
        vatAmount: vatAmount,
        brutto: totalBrutto
      },
      paymentSchedule: {
        type: 'split_50_50',
        tranche1Netto: tranche1Netto,
        tranche1Brutto: tranche1Brutto,
        tranche1Description: '50% zaliczki przed rozpoczęciem prac',
        tranche2Netto: tranche2Netto,
        tranche2Brutto: tranche2Brutto,
        tranche2Description: '50% po dostarczeniu gotowego raportu'
      },
      delivery: {
        workingDaysNumeric: timeline.numeric,
        workingDaysLabel: timeline.label,
        baseWorkingDays: pkg.workingDays,
        express: express,
        evidencePages: evidencePages,
        evidenceLimitPages: pkg.evidenceLimitPages,
        workshopMinutes: totalWorkshopMinutes,
        workshopLabel: formatWorkshop(totalWorkshopMinutes),
        includedRevisionRounds: OFFER_DATA.meta.includedRevisionRounds,
        revisionDaysLimit: OFFER_DATA.meta.revisionDaysLimit
      },
      warnings: warnings,
      notes: notes
    };
  }

  function evaluateSubscriptionPlan(plan, params) {
    var contract = OFFER_DATA.SUBSCRIPTION_CONTRACTS[params.contractType] || OFFER_DATA.SUBSCRIPTION_CONTRACTS.annual_monthly;
    var brandsCount = params.brandsCount;
    var totalUnits = params.totalUnits;
    var urgent6h = params.urgent6h;

    var brandsSupported = brandsCount <= plan.maxBrands;
    var urgentSupported = !urgent6h || plan.urgent6hFee !== null;
    var eligible = brandsSupported && urgentSupported;

    var extraBrandFee = 0;
    if (plan.id === 'pro' && brandsCount >= 3) {
      extraBrandFee = plan.extraBrandFee;
    }

    var urgent6hFee = (plan.id === 'agency' && urgent6h) ? plan.urgent6hFee : 0;
    var baseMonthlyRecurring = round2(plan.monthlyNetto + extraBrandFee + urgent6hFee);
    var adjustedMonthlyRecurring = round2(baseMonthlyRecurring * contract.multiplier);

    var overageUnits = round2(Math.max(0, totalUnits - plan.unitsPool));
    var overageMonthlyNetto = round2(overageUnits * plan.overageRate);

    var maxRolloverUnits = round2(plan.unitsPool * plan.rolloverPct);
    var unusedUnits = round2(Math.max(0, plan.unitsPool - totalUnits));
    var rolloverUnits = round2(Math.min(unusedUnits, maxRolloverUnits));
    var expiredUnits = round2(Math.max(0, unusedUnits - rolloverUnits));

    var effectiveMonthlyNetto = round2(adjustedMonthlyRecurring + overageMonthlyNetto);
    var contractRecurringTotalNetto = round2(adjustedMonthlyRecurring * contract.months);
    var contractOverageTotalNetto = round2(overageMonthlyNetto * contract.months);
    var contractTotalNetto = round2(contractRecurringTotalNetto + contractOverageTotalNetto);

    var upfrontSavingsNetto = contract.id === 'annual_upfront'
      ? round2(baseMonthlyRecurring * contract.months * contract.discountPct)
      : 0;
    var pilotConversionCreditNetto = contract.id === 'pilot_3m'
      ? round2(baseMonthlyRecurring * contract.months * contract.surchargePct)
      : 0;

    return {
      plan: plan,
      planId: plan.id,
      planName: plan.name,
      eligible: eligible,
      brandsSupported: brandsSupported,
      urgentSupported: urgentSupported,
      basePlanNetto: plan.monthlyNetto,
      extraBrandFee: extraBrandFee,
      urgent6hFee: urgent6hFee,
      baseMonthlyRecurring: baseMonthlyRecurring,
      adjustedMonthlyRecurring: adjustedMonthlyRecurring,
      unitsPool: plan.unitsPool,
      totalUnits: totalUnits,
      overageUnits: overageUnits,
      overageMonthlyNetto: overageMonthlyNetto,
      maxRolloverUnits: maxRolloverUnits,
      unusedUnits: unusedUnits,
      rolloverUnits: rolloverUnits,
      expiredUnits: expiredUnits,
      effectiveMonthlyNetto: effectiveMonthlyNetto,
      contractMonths: contract.months,
      contractRecurringTotalNetto: contractRecurringTotalNetto,
      contractOverageTotalNetto: contractOverageTotalNetto,
      contractTotalNetto: contractTotalNetto,
      upfrontSavingsNetto: upfrontSavingsNetto,
      pilotConversionCreditNetto: pilotConversionCreditNetto
    };
  }

  function calculateSubscription(rawInput) {
    var input = rawInput || {};
    var vatRate = input.vatRate !== undefined ? toNum(input.vatRate, OFFER_DATA.meta.defaultVatRate) : OFFER_DATA.meta.defaultVatRate;
    var planId = input.planId || 'auto';
    var contractType = OFFER_DATA.SUBSCRIPTION_CONTRACTS[input.contractType] ? input.contractType : 'annual_monthly';
    var contract = OFFER_DATA.SUBSCRIPTION_CONTRACTS[contractType];
    var brandsCount = Math.max(1, toInt(input.brandsCount, 1));
    var urgent6h = Boolean(input.urgent6h);

    var unitsInput = input.units || {};
    var postBannerQty = toInt(unitsInput.postBanner, 0);
    var resizeVariantQty = toInt(unitsInput.resizeVariant, 0);
    var packagingArticleQty = toInt(unitsInput.packagingArticle, 0);
    var campaignPresentationQty = toInt(unitsInput.campaignPresentation, 0);

    var unitsBreakdown = [
      {
        id: 'postBanner',
        label: OFFER_DATA.UNIT_TYPES[0].label,
        qty: postBannerQty,
        weight: 1,
        units: round2(postBannerQty * 1)
      },
      {
        id: 'resizeVariant',
        label: OFFER_DATA.UNIT_TYPES[1].label,
        qty: resizeVariantQty,
        weight: 0.25,
        units: round2(resizeVariantQty * 0.25)
      },
      {
        id: 'packagingArticle',
        label: OFFER_DATA.UNIT_TYPES[2].label,
        qty: packagingArticleQty,
        weight: 3,
        units: round2(packagingArticleQty * 3)
      },
      {
        id: 'campaignPresentation',
        label: OFFER_DATA.UNIT_TYPES[3].label,
        qty: campaignPresentationQty,
        weight: 6,
        units: round2(campaignPresentationQty * 6)
      }
    ];

    var totalUnits = round2(
      unitsBreakdown[0].units +
      unitsBreakdown[1].units +
      unitsBreakdown[2].units +
      unitsBreakdown[3].units
    );

    var comparisons = OFFER_DATA.SUBSCRIPTION_PLANS.map(function (plan) {
      return evaluateSubscriptionPlan(plan, {
        contractType: contractType,
        brandsCount: brandsCount,
        totalUnits: totalUnits,
        urgent6h: urgent6h
      });
    });

    var warnings = [];
    var notes = [];
    var selectedEval = null;

    if (planId === 'auto') {
      var eligiblePlans = comparisons.filter(function (c) { return c.eligible; });
      if (eligiblePlans.length > 0) {
        selectedEval = eligiblePlans[0];
        for (var i = 1; i < eligiblePlans.length; i++) {
          if (eligiblePlans[i].effectiveMonthlyNetto < selectedEval.effectiveMonthlyNetto) {
            selectedEval = eligiblePlans[i];
          }
        }
      } else {
        selectedEval = comparisons[comparisons.length - 1];
      }
    } else {
      selectedEval = comparisons.find(function (c) { return c.planId === planId; }) || comparisons[0];
    }

    if (brandsCount > selectedEval.plan.maxBrands) {
      warnings.push({
        severity: 'warning',
        code: 'BRANDS_LIMIT_EXCEEDED',
        message: 'Liczba marek (' + brandsCount + ') przekracza limit planu ' + selectedEval.planName + ' (maks. ' + selectedEval.plan.maxBrands + '). Wymagana indywidualna wycena lub wyższy plan.'
      });
    }

    if (urgent6h && selectedEval.plan.urgent6hFee === null) {
      warnings.push({
        severity: 'warning',
        code: 'URGENT_ONLY_AGENCY',
        message: 'Tryb pilny 6 godzin (+6 000 zł/mies., do 6 materiałów) dostępny jest wyłącznie w planie Agency.'
      });
    }

    var websiteAuditPackageId = input.websiteAuditPackage || null;
    var extraDiffScans = toInt(input.extraDiffScans, 0);
    var outOfPoolCompact = toInt(input.outOfPoolCompact, 0);
    var outOfPoolStandard = toInt(input.outOfPoolStandard, 0);

    var lineItems = [];
    lineItems.push({
      id: 'subscription_base',
      label: 'Abonament miesięczny ' + selectedEval.planName,
      detail: 'Pula ' + selectedEval.unitsPool + ' jednostek/mies., czas reakcji ' + selectedEval.plan.responseDays + ' ' + (selectedEval.plan.responseDays === 1 ? 'dzień roboczy' : 'dni robocze'),
      quantity: 1,
      unitPriceNetto: selectedEval.basePlanNetto,
      totalNetto: selectedEval.basePlanNetto
    });

    if (selectedEval.extraBrandFee > 0) {
      lineItems.push({
        id: 'extra_brand',
        label: 'Trzecia marka w planie Pro',
        detail: '+3 000 zł netto / miesiąc (powyżej 2 marek w cenie)',
        quantity: 1,
        unitPriceNetto: selectedEval.extraBrandFee,
        totalNetto: selectedEval.extraBrandFee
      });
    }

    if (selectedEval.urgent6hFee > 0) {
      lineItems.push({
        id: 'urgent_6h',
        label: 'Tryb pilny 6 godzin (Agency, do 6 materiałów/mies.)',
        detail: '+6 000 zł netto / miesiąc',
        quantity: 1,
        unitPriceNetto: selectedEval.urgent6hFee,
        totalNetto: selectedEval.urgent6hFee
      });
    }

    if (contract.id === 'annual_upfront') {
      var monthlyDiscount = round2(selectedEval.baseMonthlyRecurring - selectedEval.adjustedMonthlyRecurring);
      lineItems.push({
        id: 'contract_discount',
        label: 'Rabat za płatność jednorazową za cały rok z góry (-8%)',
        detail: 'Oszczędność w skali 12 miesięcy: ' + formatPLN(selectedEval.upfrontSavingsNetto),
        quantity: 1,
        unitPriceNetto: -monthlyDiscount,
        totalNetto: -monthlyDiscount
      });
    } else if (contract.id === 'pilot_3m') {
      var monthlySurcharge = round2(selectedEval.adjustedMonthlyRecurring - selectedEval.baseMonthlyRecurring);
      lineItems.push({
        id: 'pilot_surcharge',
        label: 'Dopłata programu pilotażowego 3 mies. (+20%)',
        detail: 'Nadpłata za 3 mies. (' + formatPLN(selectedEval.pilotConversionCreditNetto) + ') zaliczana na 1. fakturę umowy rocznej przy decyzji w 14 dni.',
        quantity: 1,
        unitPriceNetto: monthlySurcharge,
        totalNetto: monthlySurcharge
      });
    }

    if (selectedEval.overageUnits > 0) {
      lineItems.push({
        id: 'overage_units',
        label: 'Szacowana nadwyżka jednostek ponad pulę',
        detail: formatUnits(selectedEval.overageUnits) + ' jedn. × ' + formatPLN(selectedEval.plan.overageRate) + ' (rozliczenie zbiorcze na koniec miesiąca, termin 14 dni)',
        quantity: selectedEval.overageUnits,
        unitPriceNetto: selectedEval.plan.overageRate,
        totalNetto: selectedEval.overageMonthlyNetto
      });
    }

    var oneOffExtrasItems = [];
    var websiteAuditNetto = 0;
    var websiteDiscountPct = 0;

    if (websiteAuditPackageId) {
      var webPkg = getPackageById(websiteAuditPackageId);
      if (webPkg && webPkg.category === 'website') {
        var isAnnual = contract.id === 'annual_monthly' || contract.id === 'annual_upfront';
        websiteDiscountPct = isAnnual ? selectedEval.plan.websiteDiscountAnnual : selectedEval.plan.websiteDiscountPilot;
        websiteAuditNetto = round2(webPkg.priceNetto * (1 - websiteDiscountPct));

        if (!isAnnual && selectedEval.plan.websiteDiscountAnnual > 0 && websiteDiscountPct === 0) {
          notes.push('Rabat ' + Math.round(selectedEval.plan.websiteDiscountAnnual * 100) + '% na audyt strony WWW w planie ' + selectedEval.planName + ' przysługuje przy umowie rocznej (w pilotażu naliczono stawkę standardową).');
        }

        oneOffExtrasItems.push({
          id: 'website_audit',
          label: webPkg.name + (websiteDiscountPct > 0 ? ' (rabat -' + Math.round(websiteDiscountPct * 100) + '%)' : ''),
          detail: 'Cena cennikowa ' + formatPLN(webPkg.priceNetto) + (websiteDiscountPct > 0 ? ', po rabacie planu ' + selectedEval.planName : ''),
          quantity: 1,
          unitPriceNetto: websiteAuditNetto,
          totalNetto: websiteAuditNetto
        });

        if (selectedEval.planId === 'pro' && isAnnual) {
          notes.push('Po pełnym audycie strony WWW w planie Pro przysługują 2 bezpłatne skany różnicowe w roku (do 15 podstron).');
        }
      }
    }

    var diffScansNetto = 0;
    if (extraDiffScans > 0) {
      if (selectedEval.planId === 'pro') {
        diffScansNetto = round2(extraDiffScans * selectedEval.plan.extraDiffScanPrice);
        oneOffExtrasItems.push({
          id: 'extra_diff_scans',
          label: 'Dodatkowe skany różnicowe strony (Pro)',
          detail: extraDiffScans + ' szt. × ' + formatPLN(selectedEval.plan.extraDiffScanPrice),
          quantity: extraDiffScans,
          unitPriceNetto: selectedEval.plan.extraDiffScanPrice,
          totalNetto: diffScansNetto
        });
      } else {
        warnings.push({
          severity: 'info',
          code: 'DIFF_SCANS_ONLY_PRO',
          message: 'Skany różnicowe strony (2 w roku w cenie, kolejny 5 500 zł) są pozycją zdefiniowaną w planie Pro.'
        });
      }
    }

    var outOfPoolNetto = 0;
    if (outOfPoolCompact > 0) {
      var oopCompactTotal = round2(outOfPoolCompact * OFFER_DATA.OUT_OF_POOL.compactPrice);
      outOfPoolNetto = round2(outOfPoolNetto + oopCompactTotal);
      oneOffExtrasItems.push({
        id: 'out_of_pool_compact',
        label: 'Publikacje poza pulą — poziom Compact',
        detail: outOfPoolCompact + ' szt. × ' + formatPLN(OFFER_DATA.OUT_OF_POOL.compactPrice) + ' (sekcja >3 podstron, katalog, ebook lub raport >10 str.)',
        quantity: outOfPoolCompact,
        unitPriceNetto: OFFER_DATA.OUT_OF_POOL.compactPrice,
        totalNetto: oopCompactTotal
      });
    }

    if (outOfPoolStandard > 0) {
      var oopStandardTotal = round2(outOfPoolStandard * OFFER_DATA.OUT_OF_POOL.standardPrice);
      outOfPoolNetto = round2(outOfPoolNetto + oopStandardTotal);
      oneOffExtrasItems.push({
        id: 'out_of_pool_standard',
        label: 'Publikacje poza pulą — poziom Standard',
        detail: outOfPoolStandard + ' szt. × ' + formatPLN(OFFER_DATA.OUT_OF_POOL.standardPrice) + ' (rozbudowany katalog, ebook lub raport >10 str.)',
        quantity: outOfPoolStandard,
        unitPriceNetto: OFFER_DATA.OUT_OF_POOL.standardPrice,
        totalNetto: oopStandardTotal
      });
    }

    var oneOffExtrasNetto = round2(websiteAuditNetto + diffScansNetto + outOfPoolNetto);
    var monthlyVat = round2(selectedEval.effectiveMonthlyNetto * vatRate);
    var monthlyBrutto = round2(selectedEval.effectiveMonthlyNetto + monthlyVat);

    var contractVat = round2(selectedEval.contractTotalNetto * vatRate);
    var contractBrutto = round2(selectedEval.contractTotalNetto + contractVat);

    var extrasVat = round2(oneOffExtrasNetto * vatRate);
    var extrasBrutto = round2(oneOffExtrasNetto + extrasVat);

    var grandPeriodNetto = round2(selectedEval.contractTotalNetto + oneOffExtrasNetto);
    var grandPeriodVat = round2(grandPeriodNetto * vatRate);
    var grandPeriodBrutto = round2(grandPeriodNetto + grandPeriodVat);

    return {
      kind: 'subscription',
      planSelectionMode: planId,
      selectedPlan: selectedEval,
      contract: contract,
      brandsCount: brandsCount,
      urgent6h: urgent6h,
      unitsBreakdown: unitsBreakdown,
      totalUnits: totalUnits,
      comparisons: comparisons,
      lineItems: lineItems,
      oneOffExtrasItems: oneOffExtrasItems,
      websiteDiscountPct: websiteDiscountPct,
      totals: {
        monthlyNetto: selectedEval.effectiveMonthlyNetto,
        monthlyVat: monthlyVat,
        monthlyBrutto: monthlyBrutto,
        monthlyRecurringNetto: selectedEval.adjustedMonthlyRecurring,
        monthlyOverageNetto: selectedEval.overageMonthlyNetto,
        contractMonths: contract.months,
        contractNetto: selectedEval.contractTotalNetto,
        contractVat: contractVat,
        contractBrutto: contractBrutto,
        oneOffExtrasNetto: oneOffExtrasNetto,
        oneOffExtrasVat: extrasVat,
        oneOffExtrasBrutto: extrasBrutto,
        grandPeriodNetto: grandPeriodNetto,
        grandPeriodVat: grandPeriodVat,
        grandPeriodBrutto: grandPeriodBrutto,
        vatRate: vatRate
      },
      warnings: warnings,
      notes: notes
    };
  }

  function formatQuoteText(result, clientMeta) {
    var meta = clientMeta || {};
    var lines = [];
    lines.push('WYCENA AUDYTU GREENWASHINGU — ESG INSTITUTE');
    lines.push('Podstawa: Oferta handlowa z ' + OFFER_DATA.meta.offerDate + ' (ważna do ' + OFFER_DATA.meta.validUntil + ')');
    if (meta.clientName || meta.projectName || meta.contactPerson) {
      lines.push('');
      if (meta.clientName) lines.push('Klient: ' + meta.clientName);
      if (meta.projectName) lines.push('Zakres / projekt: ' + meta.projectName);
      if (meta.contactPerson) lines.push('Osoba kontaktowa: ' + meta.contactPerson);
    }

    lines.push('');
    if (result.kind === 'one_off') {
      lines.push('Wybrany pakiet: ' + result.selectedPackage.name);
      lines.push('Termin realizacji: ' + result.delivery.workingDaysLabel + (result.delivery.express ? ' (tryb Ekspres)' : ''));
      lines.push('Limit dokumentacji dowodowej: do ' + result.delivery.evidenceLimitPages + ' stron' + (result.delivery.evidencePages > 0 ? ' (zgłoszono: ' + result.delivery.evidencePages + ' str.)' : ''));
      lines.push('Warsztat z ekspertem: ' + result.delivery.workshopLabel);
      lines.push('');
      lines.push('Pozycje wyceny (netto):');
      result.lineItems.forEach(function (item) {
        lines.push('- ' + item.label + ': ' + formatPLN(item.totalNetto));
      });
      lines.push('');
      lines.push('RAZEM NETTO: ' + formatPLN(result.totals.netto));
      lines.push('VAT (' + Math.round(result.totals.vatRate * 100) + '%): ' + formatPLN(result.totals.vatAmount));
      lines.push('RAZEM BRUTTO: ' + formatPLN(result.totals.brutto));
      lines.push('');
      lines.push('Warunki płatności:');
      lines.push('- I transza (50% przed rozpoczęciem prac): ' + formatPLN(result.paymentSchedule.tranche1Netto) + ' netto (' + formatPLN(result.paymentSchedule.tranche1Brutto) + ' brutto)');
      lines.push('- II transza (50% po dostarczeniu raportu): ' + formatPLN(result.paymentSchedule.tranche2Netto) + ' netto (' + formatPLN(result.paymentSchedule.tranche2Brutto) + ' brutto)');
    } else {
      var sel = result.selectedPlan;
      lines.push('Wybrany plan abonamentowy: ' + sel.planName + ' (' + result.contract.name + ')');
      lines.push('Pula jednostek: ' + sel.unitsPool + ' jedn./mies. | Szacowane zapotrzebowanie: ' + formatUnits(result.totalUnits) + ' jedn./mies.');
      lines.push('Czas reakcji: ' + sel.plan.responseDays + ' ' + (sel.plan.responseDays === 1 ? 'dzień roboczy' : 'dni robocze') + ' | Liczba marek: ' + result.brandsCount);
      lines.push('');
      lines.push('Rozliczenie miesięczne (netto):');
      result.lineItems.forEach(function (item) {
        lines.push('- ' + item.label + ': ' + formatPLN(item.totalNetto));
      });
      lines.push('SZACOWANY KOSZT MIESIĘCZNY NETTO: ' + formatPLN(result.totals.monthlyNetto) + ' (' + formatPLN(result.totals.monthlyBrutto) + ' brutto)');
      lines.push('ŁĄCZNIE ZA OKRES (' + result.totals.contractMonths + ' mies.): ' + formatPLN(result.totals.contractNetto) + ' netto (' + formatPLN(result.totals.contractBrutto) + ' brutto)');
      if (sel.upfrontSavingsNetto > 0) {
        lines.push('Oszczędność przy płatności za rok z góry (-8%): ' + formatPLN(sel.upfrontSavingsNetto) + ' netto');
      }
      if (sel.pilotConversionCreditNetto > 0) {
        lines.push('Nadpłata pilotażu do zaliczenia na 1. fakturę umowy rocznej: ' + formatPLN(sel.pilotConversionCreditNetto) + ' netto');
      }
      if (result.oneOffExtrasItems.length > 0) {
        lines.push('');
        lines.push('Dodatkowe audyty / materiały poza pulą (jednorazowo, netto):');
        result.oneOffExtrasItems.forEach(function (item) {
          lines.push('- ' + item.label + ': ' + formatPLN(item.totalNetto));
        });
        lines.push('RAZEM POZAPULOWE NETTO: ' + formatPLN(result.totals.oneOffExtrasNetto) + ' (' + formatPLN(result.totals.oneOffExtrasBrutto) + ' brutto)');
      }
    }

    if (result.warnings.length > 0) {
      lines.push('');
      lines.push('Uwagi i zastrzeżenia:');
      result.warnings.forEach(function (w) {
        lines.push('! ' + w.message);
      });
    }

    lines.push('');
    lines.push('Kontakt: ' + OFFER_DATA.meta.contactName + ' | ' + OFFER_DATA.meta.contactEmail + ' | ' + OFFER_DATA.meta.contactPhone);
    return lines.join('\n');
  }

  return {
    OFFER_DATA: OFFER_DATA,
    calculateOneOff: calculateOneOff,
    calculateSubscription: calculateSubscription,
    formatQuoteText: formatQuoteText,
    formatPLN: formatPLN,
    formatUnits: formatUnits,
    getPackageById: getPackageById,
    getPlanById: getPlanById
  };
});
