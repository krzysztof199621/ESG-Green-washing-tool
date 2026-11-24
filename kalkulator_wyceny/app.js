(function () {
  'use strict';

  var P = window.ESGPricing;
  if (!P) return;

  var state = {
    activeTab: 'one-off',
    oneOff: {
      mode: 'scope',
      scopeType: 'material',
      hasCampaignPage: false,
      creatives: 6,
      packagings: 0,
      sites: 1,
      subpages: 8,
      packageId: 'przeglad',
      extraCreatives: 0,
      extraPackagings: 0,
      extraSites: 0,
      evidencePages: 25,
      dictionary: false,
      english: false,
      express: false,
      extraWorkshopHours: 0,
      extraRevisionHours: 0,
      subscriptionDiscountPct: 0
    },
    subscription: {
      planId: 'auto',
      contractType: 'annual_monthly',
      brandsCount: 1,
      urgent6h: false,
      units: {
        postBanner: 4,
        resizeVariant: 4,
        packagingArticle: 1,
        campaignPresentation: 0
      },
      websiteAuditPackage: '',
      extraDiffScans: 0,
      outOfPoolCompact: 0,
      outOfPoolStandard: 0
    }
  };

  function el(id) {
    return document.getElementById(id);
  }

  function escapeHtml(text) {
    return String(text == null ? '' : text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getVatRate() {
    var v = Number(el('vatRateSelect').value);
    return Number.isFinite(v) ? v : 0.23;
  }

  function getClientMeta() {
    return {
      clientName: el('clientName').value.trim(),
      projectName: el('projectName').value.trim(),
      contactPerson: el('contactPerson').value.trim()
    };
  }

  function syncInputsFromState() {
    el('modeScopeBtn').classList.toggle('active', state.oneOff.mode === 'scope');
    el('modeManualBtn').classList.toggle('active', state.oneOff.mode === 'manual');
    el('scopeModeSection').classList.toggle('hidden', state.oneOff.mode !== 'scope');
    el('manualModeSection').classList.toggle('hidden', state.oneOff.mode !== 'manual');

    el('scopeMaterialBtn').classList.toggle('active', state.oneOff.scopeType === 'material');
    el('scopeWebsiteBtn').classList.toggle('active', state.oneOff.scopeType === 'website');
    el('scopeMaterialInputs').classList.toggle('hidden', state.oneOff.scopeType !== 'material');
    el('scopeWebsiteInputs').classList.toggle('hidden', state.oneOff.scopeType !== 'website');

    el('hasCampaignPage').checked = state.oneOff.hasCampaignPage;
    el('creativesCount').value = state.oneOff.creatives;
    el('packagingsCount').value = state.oneOff.packagings;
    el('sitesCount').value = state.oneOff.sites;
    el('subpagesCount').value = state.oneOff.subpages;

    el('manualPackageSelect').value = state.oneOff.packageId;
    el('manualExtraCreatives').value = state.oneOff.extraCreatives;
    el('manualExtraPackagings').value = state.oneOff.extraPackagings;
    el('manualExtraSites').value = state.oneOff.extraSites;

    el('evidencePages').value = state.oneOff.evidencePages;
    el('optDictionary').checked = state.oneOff.dictionary;
    el('optEnglish').checked = state.oneOff.english;
    el('optExpress').checked = state.oneOff.express;
    el('extraWorkshopHours').value = state.oneOff.extraWorkshopHours;
    el('extraRevisionHours').value = state.oneOff.extraRevisionHours;
    el('subDiscountSelect').value = String(state.oneOff.subscriptionDiscountPct);

    // Subscription inputs
    var contractButtons = el('contractTypeGroup').querySelectorAll('button[data-contract]');
    for (var i = 0; i < contractButtons.length; i++) {
      var btn = contractButtons[i];
      btn.classList.toggle('active', btn.getAttribute('data-contract') === state.subscription.contractType);
    }

    el('subPlanSelect').value = state.subscription.planId;
    el('brandsCount').value = state.subscription.brandsCount;
    el('optUrgent6h').checked = state.subscription.urgent6h;
    el('unitPostBanner').value = state.subscription.units.postBanner;
    el('unitResizeVariant').value = state.subscription.units.resizeVariant;
    el('unitPackagingArticle').value = state.subscription.units.packagingArticle;
    el('unitCampaignPresentation').value = state.subscription.units.campaignPresentation;
    el('subWebsitePackage').value = state.subscription.websiteAuditPackage;
    el('subExtraDiffScans').value = state.subscription.extraDiffScans;
    el('outOfPoolCompact').value = state.subscription.outOfPoolCompact;
    el('outOfPoolStandard').value = state.subscription.outOfPoolStandard;
  }

  function renderTabs() {
    var tabButtons = document.querySelectorAll('.tab-btn');
    for (var i = 0; i < tabButtons.length; i++) {
      var btn = tabButtons[i];
      var tabId = btn.getAttribute('data-tab');
      var active = tabId === state.activeTab;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    }

    var panels = document.querySelectorAll('.tab-panel');
    for (var j = 0; j < panels.length; j++) {
      panels[j].classList.toggle('active', panels[j].id === 'tab-' + state.activeTab);
    }

    var isReference = state.activeTab === 'reference';
    el('workspaceLayout').classList.toggle('full-width', isReference);
    el('summaryAside').classList.toggle('hidden', isReference);
  }

  function renderPrintClientAndExclusions() {
    var meta = getClientMeta();
    var parts = [];
    if (meta.clientName) parts.push('<div><strong>Klient:</strong> ' + escapeHtml(meta.clientName) + '</div>');
    if (meta.projectName) parts.push('<div><strong>Projekt / zakres:</strong> ' + escapeHtml(meta.projectName) + '</div>');
    if (meta.contactPerson) parts.push('<div><strong>Osoba kontaktowa:</strong> ' + escapeHtml(meta.contactPerson) + '</div>');
    parts.push('<div><strong>Data wygenerowania:</strong> ' + new Date().toLocaleDateString('pl-PL') + '</div>');
    el('printClientBox').innerHTML = parts.join('');

    el('printExclusionsList').innerHTML = P.OFFER_DATA.EXCLUSIONS.map(function (item) {
      return '<li>' + escapeHtml(item) + '</li>';
    }).join('');
  }

  function renderOneOffSummary(res) {
    var vatPct = Math.round(res.totals.vatRate * 100);
    el('summaryKicker').textContent = res.mode === 'scope'
      ? 'Audyt jednorazowy — dobór automatyczny'
      : 'Audyt jednorazowy — wybór ręczny';
    el('summaryPackageName').textContent = res.selectedPackage.name;
    el('summaryMainPrice').textContent = P.formatPLN(res.totals.netto);
    el('summaryMainUnit').textContent = 'netto';
    el('summaryBruttoPrice').textContent = P.formatPLN(res.totals.brutto) + ' brutto (VAT ' + vatPct + '%)';
    el('summarySubNote').textContent = res.selectedPackage.includedSummary;

    // Evidence meter update
    var limit = res.delivery.evidenceLimitPages;
    var pages = res.delivery.evidencePages;
    var pct = Math.min(100, Math.round((pages / Math.max(1, limit)) * 100));
    el('evidenceMeterLabel').textContent = pages + ' stron';
    el('evidenceLimitLabel').textContent = 'Limit pakietu: ' + limit + ' stron';
    var fill = el('evidenceMeterFill');
    fill.style.width = pct + '%';
    fill.classList.toggle('over', pages > limit);
    fill.classList.toggle('ok', pages <= limit);

    // Stats
    el('summaryStats').innerHTML = [
      '<div class="stat-box"><span>Termin realizacji</span><strong>' + escapeHtml(res.delivery.workingDaysLabel) + (res.delivery.express ? ' (Ekspres)' : '') + '</strong></div>',
      '<div class="stat-box"><span>Limit dowodów</span><strong>do ' + limit + ' str.</strong></div>',
      '<div class="stat-box"><span>Warsztat</span><strong>' + escapeHtml(res.delivery.workshopLabel) + '</strong></div>'
    ].join('');

    // Lines
    el('summaryLinesBody').innerHTML = res.lineItems.map(function (item) {
      var cls = item.totalNetto < 0 ? ' class="discount"' : '';
      return '<tr' + cls + '>' +
        '<td>' + escapeHtml(item.label) + '<span class="line-detail">' + escapeHtml(item.detail) + '</span></td>' +
        '<td>' + P.formatPLN(item.totalNetto) + '</td>' +
        '</tr>';
    }).join('');

    // Totals
    el('summaryTotalsBox').innerHTML =
      '<div class="totals-row"><span>Suma netto</span><strong>' + P.formatPLN(res.totals.netto) + '</strong></div>' +
      '<div class="totals-row"><span>Podatek VAT (' + vatPct + '%)</span><span>' + P.formatPLN(res.totals.vatAmount) + '</span></div>' +
      '<div class="totals-row grand"><span>Razem brutto</span><span>' + P.formatPLN(res.totals.brutto) + '</span></div>';

    // Payment schedule
    el('summaryScheduleBox').innerHTML =
      '<strong>Harmonogram rozliczenia (2 równe transze)</strong>' +
      '<ul>' +
        '<li>I transza (50% przed startem): <strong>' + P.formatPLN(res.paymentSchedule.tranche1Netto) + '</strong> netto (' + P.formatPLN(res.paymentSchedule.tranche1Brutto) + ' brutto)</li>' +
        '<li>II transza (50% po raporcie): <strong>' + P.formatPLN(res.paymentSchedule.tranche2Netto) + '</strong> netto (' + P.formatPLN(res.paymentSchedule.tranche2Brutto) + ' brutto)</li>' +
        '<li>1 runda uwag redakcyjnych w ciągu 5 dni rob. od raportu w cenie.</li>' +
      '</ul>';

    // Alerts & notes
    var alertsHtml = [];
    res.warnings.forEach(function (w) {
      alertsHtml.push('<div class="alert ' + escapeHtml(w.severity) + '">' + escapeHtml(w.message) + '</div>');
    });
    res.notes.forEach(function (n) {
      alertsHtml.push('<div class="alert info">' + escapeHtml(n) + '</div>');
    });
    el('summaryAlerts').innerHTML = alertsHtml.join('');

    // Alternatives card
    var altCard = el('oneOffAlternativesCard');
    if (res.mode === 'scope' && res.alternatives.length > 1) {
      altCard.classList.remove('hidden');
      el('oneOffAlternativesGrid').innerHTML = res.alternatives.map(function (alt) {
        var isSel = alt.packageId === res.selectedPackage.id;
        return '<button type="button" class="alt-card' + (isSel ? ' selected' : '') + '" data-choose-pkg="' + escapeHtml(alt.packageId) + '" data-extra-c="' + alt.extraCreatives + '" data-extra-p="' + alt.extraPackagings + '" data-extra-s="' + alt.extraSites + '">' +
          '<div class="alt-top"><strong>' + escapeHtml(alt.shortName) + '</strong>' + (isSel ? '<span class="alt-tag">Najkorzystniejszy</span>' : '') + '</div>' +
          '<div class="alt-price">' + P.formatPLN(alt.scopeNetto) + ' netto</div>' +
          '<div class="alt-meta">Termin: ' + alt.workingDays + ' dni rob. · Dowody: ' + alt.evidenceLimitPages + ' str.</div>' +
          '</button>';
      }).join('');
    } else {
      altCard.classList.add('hidden');
    }
  }

  function renderSubscriptionSummary(res) {
    var sel = res.selectedPlan;
    var vatPct = Math.round(res.totals.vatRate * 100);

    // Update unit subtotals
    el('subTotalPostBanner').textContent = P.formatUnits(res.unitsBreakdown[0].units) + ' jedn.';
    el('subTotalResizeVariant').textContent = P.formatUnits(res.unitsBreakdown[1].units) + ' jedn.';
    el('subTotalPackagingArticle').textContent = P.formatUnits(res.unitsBreakdown[2].units) + ' jedn.';
    el('subTotalCampaignPresentation').textContent = P.formatUnits(res.unitsBreakdown[3].units) + ' jedn.';

    el('unitsPoolSummary').textContent = 'Zapotrzebowanie: ' + P.formatUnits(res.totalUnits) + ' jedn. / pula planu ' + sel.planName + ': ' + sel.unitsPool + ' jedn.';
    el('unitsPoolHint').textContent = sel.overageUnits > 0
      ? 'Nadwyżka: ' + P.formatUnits(sel.overageUnits) + ' jedn. × ' + P.formatPLN(sel.plan.overageRate)
      : 'Przeniesienie na kolejny mies.: ' + P.formatUnits(sel.rolloverUnits) + ' jedn. (maks. ' + P.formatUnits(sel.maxRolloverUnits) + ')';

    var fillPct = Math.min(100, Math.round((res.totalUnits / Math.max(1, sel.unitsPool)) * 100));
    var meterFill = el('unitsMeterFill');
    meterFill.style.width = fillPct + '%';
    meterFill.classList.toggle('over', sel.overageUnits > 0);
    meterFill.classList.toggle('ok', sel.overageUnits === 0);

    // Diff scans visibility / hint
    el('subDiffScansWrap').style.opacity = sel.planId === 'pro' ? '1' : '0.6';

    // Hero summary
    el('summaryKicker').textContent = 'Abonament ' + res.contract.name;
    el('summaryPackageName').textContent = 'Plan ' + sel.planName + (res.planSelectionMode === 'auto' ? ' (rekomendowany)' : '');
    el('summaryMainPrice').textContent = P.formatPLN(res.totals.monthlyNetto);
    el('summaryMainUnit').textContent = 'netto / mies.';
    el('summaryBruttoPrice').textContent = P.formatPLN(res.totals.monthlyBrutto) + ' brutto / mies. (VAT ' + vatPct + '%)';
    el('summarySubNote').textContent = 'Łącznie za ' + res.totals.contractMonths + ' mies.: ' + P.formatPLN(res.totals.contractNetto) + ' netto' +
      (res.totals.oneOffExtrasNetto > 0 ? ' + jednorazowo ' + P.formatPLN(res.totals.oneOffExtrasNetto) + ' netto' : '');

    var daysLabel = sel.plan.responseDays === 1 ? '1 dzień rob.' : sel.plan.responseDays + ' dni rob.';
    el('summaryStats').innerHTML = [
      '<div class="stat-box"><span>Pula jednostek</span><strong>' + sel.unitsPool + ' / mies.</strong></div>',
      '<div class="stat-box"><span>Czas reakcji</span><strong>' + (sel.urgent6hFee > 0 ? '6h (pilny) / 1d' : daysLabel) + '</strong></div>',
      '<div class="stat-box"><span>Przeniesienie</span><strong>do ' + P.formatUnits(sel.maxRolloverUnits) + ' jedn.</strong></div>'
    ].join('');

    var allLines = res.lineItems.concat(res.oneOffExtrasItems);
    el('summaryLinesBody').innerHTML = allLines.map(function (item) {
      var cls = item.totalNetto < 0 ? ' class="discount"' : '';
      return '<tr' + cls + '>' +
        '<td>' + escapeHtml(item.label) + '<span class="line-detail">' + escapeHtml(item.detail) + '</span></td>' +
        '<td>' + P.formatPLN(item.totalNetto) + '</td>' +
        '</tr>';
    }).join('');

    var totalsRows = [
      '<div class="totals-row"><span>Miesięcznie netto (z nadwyżką)</span><strong>' + P.formatPLN(res.totals.monthlyNetto) + '</strong></div>',
      '<div class="totals-row"><span>Miesięcznie brutto (VAT ' + vatPct + '%)</span><span>' + P.formatPLN(res.totals.monthlyBrutto) + '</span></div>',
      '<div class="totals-row"><span>Okres umowy (' + res.totals.contractMonths + ' mies.) netto</span><strong>' + P.formatPLN(res.totals.contractNetto) + '</strong></div>'
    ];
    if (res.totals.oneOffExtrasNetto > 0) {
      totalsRows.push('<div class="totals-row"><span>Audyt WWW i poza pulą (jednorazowo netto)</span><strong>' + P.formatPLN(res.totals.oneOffExtrasNetto) + '</strong></div>');
    }
    totalsRows.push('<div class="totals-row grand"><span>Łącznie w okresie brutto</span><span>' + P.formatPLN(res.totals.grandPeriodBrutto) + '</span></div>');
    el('summaryTotalsBox').innerHTML = totalsRows.join('');

    // Schedule box
    var scheduleItems = [];
    if (res.contract.id === 'annual_upfront') {
      scheduleItems.push('<li>Płatność za 12 mies. z góry (część stała po rabacie -8%): <strong>' + P.formatPLN(sel.contractRecurringTotalNetto) + '</strong> netto (oszczędność ' + P.formatPLN(sel.upfrontSavingsNetto) + ' netto).</li>');
    } else if (res.contract.id === 'pilot_3m') {
      scheduleItems.push('<li>Pilotaż 3 mies. (+20%): miesięczna opłata stała <strong>' + P.formatPLN(sel.adjustedMonthlyRecurring) + '</strong> netto.</li>');
      scheduleItems.push('<li>Nadpłata <strong>' + P.formatPLN(sel.pilotConversionCreditNetto) + '</strong> netto zostanie w całości zaliczona na 1. fakturę umowy 12-mies. przy decyzji w ciągu 14 dni od zakończenia pilotażu.</li>');
    } else {
      scheduleItems.push('<li>Umowa 12 mies., stała opłata <strong>' + P.formatPLN(sel.adjustedMonthlyRecurring) + '</strong> netto płatna miesięcznie z góry.</li>');
    }
    scheduleItems.push('<li>Jednostki ponadlimitowe rozliczane zbiorczo na koniec miesiąca (termin płatności 14 dni).</li>');
    el('summaryScheduleBox').innerHTML = '<strong>Warunki rozliczenia abonamentu</strong><ul>' + scheduleItems.join('') + '</ul>';

    // Alerts
    var alertsHtml = [];
    res.warnings.forEach(function (w) {
      alertsHtml.push('<div class="alert ' + escapeHtml(w.severity) + '">' + escapeHtml(w.message) + '</div>');
    });
    res.notes.forEach(function (n) {
      alertsHtml.push('<div class="alert info">' + escapeHtml(n) + '</div>');
    });
    if (sel.plan.perks && sel.plan.perks.length > 0) {
      alertsHtml.push('<div class="alert info"><strong>W cenie planu ' + escapeHtml(sel.planName) + ':</strong> ' + escapeHtml(sel.plan.perks.join(' · ')) + '</div>');
    }
    el('summaryAlerts').innerHTML = alertsHtml.join('');

    // Comparison table
    el('subComparisonBody').innerHTML = res.comparisons.map(function (c) {
      var isSel = c.planId === sel.planId;
      var statusText = c.plan.responseDays + ' ' + (c.plan.responseDays === 1 ? 'dzień rob.' : 'dni rob.');
      if (!c.brandsSupported) statusText = 'Przekroczony limit marek (' + c.plan.maxBrands + ')';
      else if (!c.urgentSupported) statusText = 'Brak trybu pilnego 6h';
      return '<tr' + (isSel ? ' class="highlight-row"' : '') + '>' +
        '<td><strong>' + escapeHtml(c.planName) + '</strong>' + (isSel ? ' <span class="alt-tag">Wybrany</span>' : '') + '</td>' +
        '<td class="num">' + c.unitsPool + ' jedn.</td>' +
        '<td class="num">' + (c.overageUnits > 0 ? P.formatUnits(c.overageUnits) + ' jedn. (' + P.formatPLN(c.overageMonthlyNetto) + ')' : '0') + '</td>' +
        '<td class="num">' + P.formatPLN(c.adjustedMonthlyRecurring) + '</td>' +
        '<td class="num"><strong>' + P.formatPLN(c.effectiveMonthlyNetto) + '</strong></td>' +
        '<td>' + escapeHtml(statusText) + '</td>' +
        '</tr>';
    }).join('');
  }

  function getCurrentResult() {
    var vatRate = getVatRate();
    if (state.activeTab === 'subscription') {
      return P.calculateSubscription(Object.assign({}, state.subscription, { vatRate: vatRate }));
    }
    return P.calculateOneOff(Object.assign({}, state.oneOff, { vatRate: vatRate }));
  }

  function update() {
    renderTabs();
    renderPrintClientAndExclusions();

    if (state.activeTab === 'subscription') {
      var subRes = P.calculateSubscription(Object.assign({}, state.subscription, { vatRate: getVatRate() }));
      renderSubscriptionSummary(subRes);
    } else {
      var oneOffRes = P.calculateOneOff(Object.assign({}, state.oneOff, { vatRate: getVatRate() }));
      renderOneOffSummary(oneOffRes);
    }
  }

  function applyPreset(name) {
    state.activeTab = 'one-off';
    state.oneOff.mode = 'scope';
    state.oneOff.extraWorkshopHours = 0;
    state.oneOff.extraRevisionHours = 0;
    state.oneOff.subscriptionDiscountPct = 0;

    if (name === 'single-material') {
      state.oneOff.scopeType = 'material';
      state.oneOff.hasCampaignPage = false;
      state.oneOff.creatives = 0;
      state.oneOff.packagings = 1;
      state.oneOff.evidencePages = 20;
      state.oneOff.dictionary = false;
      state.oneOff.english = false;
      state.oneOff.express = false;
    } else if (name === 'campaign-en') {
      state.oneOff.scopeType = 'material';
      state.oneOff.hasCampaignPage = false;
      state.oneOff.creatives = 15;
      state.oneOff.packagings = 0;
      state.oneOff.evidencePages = 35;
      state.oneOff.dictionary = false;
      state.oneOff.english = true;
      state.oneOff.express = false;
    } else if (name === 'web-standard-dict') {
      state.oneOff.scopeType = 'website';
      state.oneOff.sites = 1;
      state.oneOff.subpages = 15;
      state.oneOff.evidencePages = 60;
      state.oneOff.dictionary = true;
      state.oneOff.english = false;
      state.oneOff.express = false;
    } else if (name === 'web-enterprise-express') {
      state.oneOff.scopeType = 'website';
      state.oneOff.sites = 2;
      state.oneOff.subpages = 30;
      state.oneOff.evidencePages = 80;
      state.oneOff.dictionary = false;
      state.oneOff.english = false;
      state.oneOff.express = true;
    }

    syncInputsFromState();
    update();
  }

  function resetAll() {
    state.oneOff = {
      mode: 'scope',
      scopeType: 'material',
      hasCampaignPage: false,
      creatives: 6,
      packagings: 0,
      sites: 1,
      subpages: 8,
      packageId: 'przeglad',
      extraCreatives: 0,
      extraPackagings: 0,
      extraSites: 0,
      evidencePages: 25,
      dictionary: false,
      english: false,
      express: false,
      extraWorkshopHours: 0,
      extraRevisionHours: 0,
      subscriptionDiscountPct: 0
    };
    state.subscription = {
      planId: 'auto',
      contractType: 'annual_monthly',
      brandsCount: 1,
      urgent6h: false,
      units: {
        postBanner: 4,
        resizeVariant: 4,
        packagingArticle: 1,
        campaignPresentation: 0
      },
      websiteAuditPackage: '',
      extraDiffScans: 0,
      outOfPoolCompact: 0,
      outOfPoolStandard: 0
    };
    el('copyFeedback').textContent = '';
    syncInputsFromState();
    update();
  }

  function bindEvents() {
    // Tabs
    var tabButtons = document.querySelectorAll('.tab-btn');
    for (var i = 0; i < tabButtons.length; i++) {
      tabButtons[i].addEventListener('click', function () {
        state.activeTab = this.getAttribute('data-tab');
        update();
      });
    }

    // Presets
    var presetBtns = document.querySelectorAll('.preset-btn');
    for (var p = 0; p < presetBtns.length; p++) {
      presetBtns[p].addEventListener('click', function () {
        applyPreset(this.getAttribute('data-preset'));
      });
    }

    // Client meta & VAT
    ['clientName', 'projectName', 'contactPerson', 'vatRateSelect'].forEach(function (id) {
      el(id).addEventListener('input', update);
      el(id).addEventListener('change', update);
    });

    // Mode & Scope toggles
    el('modeScopeBtn').addEventListener('click', function () {
      state.oneOff.mode = 'scope';
      syncInputsFromState();
      update();
    });
    el('modeManualBtn').addEventListener('click', function () {
      state.oneOff.mode = 'manual';
      syncInputsFromState();
      update();
    });
    el('scopeMaterialBtn').addEventListener('click', function () {
      state.oneOff.scopeType = 'material';
      syncInputsFromState();
      update();
    });
    el('scopeWebsiteBtn').addEventListener('click', function () {
      state.oneOff.scopeType = 'website';
      syncInputsFromState();
      update();
    });

    // One-off inputs
    var oneOffBindings = [
      ['hasCampaignPage', 'checked', 'hasCampaignPage'],
      ['creativesCount', 'int', 'creatives'],
      ['packagingsCount', 'int', 'packagings'],
      ['sitesCount', 'intMin1', 'sites'],
      ['subpagesCount', 'intMin1', 'subpages'],
      ['manualPackageSelect', 'str', 'packageId'],
      ['manualExtraCreatives', 'int', 'extraCreatives'],
      ['manualExtraPackagings', 'int', 'extraPackagings'],
      ['manualExtraSites', 'int', 'extraSites'],
      ['evidencePages', 'int', 'evidencePages'],
      ['optDictionary', 'checked', 'dictionary'],
      ['optEnglish', 'checked', 'english'],
      ['optExpress', 'checked', 'express'],
      ['extraWorkshopHours', 'num', 'extraWorkshopHours'],
      ['extraRevisionHours', 'num', 'extraRevisionHours'],
      ['subDiscountSelect', 'num', 'subscriptionDiscountPct']
    ];

    oneOffBindings.forEach(function (b) {
      var dom = el(b[0]);
      var handler = function () {
        if (b[1] === 'checked') state.oneOff[b[2]] = dom.checked;
        else if (b[1] === 'int') state.oneOff[b[2]] = Math.max(0, parseInt(dom.value || '0', 10) || 0);
        else if (b[1] === 'intMin1') state.oneOff[b[2]] = Math.max(1, parseInt(dom.value || '1', 10) || 1);
        else if (b[1] === 'num') state.oneOff[b[2]] = Math.max(0, parseFloat(dom.value || '0') || 0);
        else state.oneOff[b[2]] = dom.value;
        update();
      };
      dom.addEventListener('input', handler);
      dom.addEventListener('change', handler);
    });

    // Alternative cards click -> switch to manual mode with chosen package & extras
    el('oneOffAlternativesGrid').addEventListener('click', function (evt) {
      var card = evt.target.closest('[data-choose-pkg]');
      if (!card) return;
      state.oneOff.mode = 'manual';
      state.oneOff.packageId = card.getAttribute('data-choose-pkg');
      state.oneOff.extraCreatives = parseInt(card.getAttribute('data-extra-c') || '0', 10) || 0;
      state.oneOff.extraPackagings = parseInt(card.getAttribute('data-extra-p') || '0', 10) || 0;
      state.oneOff.extraSites = parseInt(card.getAttribute('data-extra-s') || '0', 10) || 0;
      syncInputsFromState();
      update();
    });

    // Contract buttons
    el('contractTypeGroup').addEventListener('click', function (evt) {
      var btn = evt.target.closest('[data-contract]');
      if (!btn) return;
      state.subscription.contractType = btn.getAttribute('data-contract');
      syncInputsFromState();
      update();
    });

    // Subscription bindings
    var subBindings = [
      ['subPlanSelect', 'str', 'planId'],
      ['brandsCount', 'intMin1', 'brandsCount'],
      ['optUrgent6h', 'checked', 'urgent6h'],
      ['subWebsitePackage', 'str', 'websiteAuditPackage'],
      ['subExtraDiffScans', 'int', 'extraDiffScans'],
      ['outOfPoolCompact', 'int', 'outOfPoolCompact'],
      ['outOfPoolStandard', 'int', 'outOfPoolStandard']
    ];
    subBindings.forEach(function (b) {
      var dom = el(b[0]);
      var handler = function () {
        if (b[1] === 'checked') state.subscription[b[2]] = dom.checked;
        else if (b[1] === 'int') state.subscription[b[2]] = Math.max(0, parseInt(dom.value || '0', 10) || 0);
        else if (b[1] === 'intMin1') state.subscription[b[2]] = Math.max(1, parseInt(dom.value || '1', 10) || 1);
        else state.subscription[b[2]] = dom.value;
        update();
      };
      dom.addEventListener('input', handler);
      dom.addEventListener('change', handler);
    });

    // Units bindings
    [
      ['unitPostBanner', 'postBanner'],
      ['unitResizeVariant', 'resizeVariant'],
      ['unitPackagingArticle', 'packagingArticle'],
      ['unitCampaignPresentation', 'campaignPresentation']
    ].forEach(function (u) {
      var dom = el(u[0]);
      var handler = function () {
        state.subscription.units[u[1]] = Math.max(0, parseInt(dom.value || '0', 10) || 0);
        update();
      };
      dom.addEventListener('input', handler);
      dom.addEventListener('change', handler);
    });

    // Generic stepper buttons
    var steppers = document.querySelectorAll('.stepper-btn[data-target]');
    for (var s = 0; s < steppers.length; s++) {
      steppers[s].addEventListener('click', function () {
        var targetId = this.getAttribute('data-target');
        var delta = parseFloat(this.getAttribute('data-delta') || '0');
        var input = el(targetId);
        if (!input) return;
        var min = input.min !== '' ? parseFloat(input.min) : 0;
        var current = parseFloat(input.value || '0') || 0;
        var next = Math.max(min, Math.round((current + delta) * 100) / 100);
        input.value = next;
        input.dispatchEvent(new Event('input', { bubbles: true }));
      });
    }

    // Actions
    el('copyQuoteBtn').addEventListener('click', function () {
      var text = P.formatQuoteText(getCurrentResult(), getClientMeta());
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
          el('copyFeedback').textContent = 'Skopiowano podsumowanie wyceny do schowka.';
        }, function () {
          fallbackCopy(text);
        });
      } else {
        fallbackCopy(text);
      }
    });

    el('printQuoteBtn').addEventListener('click', function () {
      window.print();
    });

    el('resetBtn').addEventListener('click', resetAll);
  }

  function fallbackCopy(text) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand('copy');
      el('copyFeedback').textContent = 'Skopiowano podsumowanie wyceny do schowka.';
    } catch (e) {
      el('copyFeedback').textContent = 'Nie udało się skopiować automatycznie — użyj opcji Drukuj / PDF.';
    }
    document.body.removeChild(ta);
  }

  syncInputsFromState();
  bindEvents();
  update();
})();
