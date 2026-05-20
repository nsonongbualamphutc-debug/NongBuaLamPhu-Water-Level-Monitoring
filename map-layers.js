/**
 * map-layers.js
 * จัดการ layers ทั้งหมดบนแผนที่ Leaflet
 * — Basemap switcher (ESRI Light/Topo/Street/NatGeo/Dark, OSM)
 * — Layer panel (toggle เส้นน้ำ / สถานี / flow / choropleth / heatmap)
 * — Choropleth overlay (เทียบ % ระดับน้ำ/ตลิ่ง รายอำเภอ)
 * — Heatmap overlay (leaflet.heat)
 * — Auto-refresh countdown badge
 *
 * ระบบสถานการณ์น้ำ จ.หนองบัวลำภู — สำนักงานสถิติจังหวัดหนองบัวลำภู
 * ต้องโหลดหลัง: Leaflet.js, leaflet.heat, config.js, waterways-loader.js
 */

(function (global) {
  'use strict';

  /* ════════════════════════════════════════════
   *  BASEMAP TILE DEFINITIONS
   * ════════════════════════════════════════════ */
  const BASEMAPS = [
    {
      id:      'esri-light',
      name:    'ESRI Light Gray',
      desc:    'สะอาด เหมาะสำหรับแดชบอร์ด',
      url:     'https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}',
      options: { attribution: '&copy; ESRI', maxZoom: 16 },
    },
    {
      id:      'esri-topo',
      name:    'ESRI Topo',
      desc:    'แสดงภูมิประเทศและเส้นชั้นความสูง',
      url:     'https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}',
      options: { attribution: '&copy; ESRI', maxZoom: 18 },
    },
    {
      id:      'esri-street',
      name:    'ESRI Street',
      desc:    'แผนที่ถนนละเอียด',
      url:     'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
      options: { attribution: '&copy; ESRI', maxZoom: 18 },
    },
    {
      id:      'esri-natgeo',
      name:    'National Geographic',
      desc:    'สไตล์ National Geographic',
      url:     'https://server.arcgisonline.com/ArcGIS/rest/services/NatGeo_World_Map/MapServer/tile/{z}/{y}/{x}',
      options: { attribution: '&copy; ESRI, Nat. Geo.', maxZoom: 16 },
    },
    {
      id:      'esri-dark',
      name:    'ESRI Dark Canvas',
      desc:    'พื้นหลังมืด เหมาะ Night mode',
      url:     'https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}',
      options: { attribution: '&copy; ESRI', maxZoom: 16 },
    },
    {
      id:      'osm',
      name:    'OpenStreetMap',
      desc:    'ข้อมูลเปิด OpenStreetMap',
      url:     'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
      options: {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OSM</a>',
        maxZoom: 19,
      },
    },
  ];

  /* ════════════════════════════════════════════
   *  CHOROPLETH CONFIG
   *  อำเภอใน จ.หนองบัวลำภู (6 อำเภอ)
   * ════════════════════════════════════════════ */
  const AMPHOE_LIST = [
    'เมืองหนองบัวลำภู', 'นากลาง', 'นาวัง',
    'ศรีบุญเรือง', 'สุวรรณคูหา', 'โนนสัง',
  ];

  function choroplethColor(pct) {
    /* 0–30: น้อย(ฟ้า) → 30–70: ปกติ(เขียว) → 70–90: เฝ้าระวัง(เหลือง) → 90+: วิกฤติ(แดง) */
    if (pct == null || isNaN(pct)) return '#e2e8f0';
    if (pct >= 95)  return '#7f1d1d';
    if (pct >= 85)  return '#dc2626';
    if (pct >= 70)  return '#f59e0b';
    if (pct >= 30)  return '#16a34a';
    return '#0ea5e9';
  }

  /* ════════════════════════════════════════════
   *  PRIVATE STATE
   * ════════════════════════════════════════════ */
  let _map       = null;
  let _baseLayers = {};      // id → L.TileLayer
  let _activeBase = null;    // ชั้นปัจจุบัน
  let _activeBaseId = 'esri-light';

  /* overlay layers */
  let _layerWaterwayLines  = null;
  let _layerStationPoints  = null;
  let _layerFlowAnim       = null;
  let _layerChoropleth     = null;
  let _layerHeatmap        = null;

  /* choropleth dimension  */
  let _choroDimension = 'water_pct'; // 'water_pct' | 'rain_24' | 'risk'

  /* live data cache */
  let _summaryData = null;

  /* refresh countdown */
  let _refreshInterval   = null;
  let _refreshCountdown  = 0;
  let _refreshPeriodMs   = 5 * 60 * 1000; // 5 นาที

  /* ════════════════════════════════════════════
   *  BASEMAP
   * ════════════════════════════════════════════ */
  function initBasemaps() {
    BASEMAPS.forEach(function (bm) {
      _baseLayers[bm.id] = L.tileLayer(bm.url, bm.options);
    });
    /* default */
    _baseLayers[_activeBaseId].addTo(_map);
    _activeBase = _baseLayers[_activeBaseId];
  }

  function switchBasemap(id) {
    if (!_baseLayers[id] || id === _activeBaseId) return;
    if (_activeBase) _map.removeLayer(_activeBase);
    _activeBase   = _baseLayers[id];
    _activeBaseId = id;
    _activeBase.addTo(_map);
    _activeBase.bringToBack();
    /* sync UI */
    document.querySelectorAll('.bm-option').forEach(function (el) {
      el.classList.toggle('active', el.dataset.bm === id);
    });
    showToast('เปลี่ยนแผนที่พื้นหลัง: ' + (BASEMAPS.find(b => b.id === id) || {}).name, 'info');
  }

  /* ════════════════════════════════════════════
   *  BASEMAP DROPDOWN UI
   * ════════════════════════════════════════════ */
  function buildBasemapDropdown() {
    const menu = document.querySelector('.basemap-menu');
    if (!menu) return;
    menu.innerHTML = BASEMAPS.map(function (bm) {
      return (
        '<div class="bm-option' + (bm.id === _activeBaseId ? ' active' : '') + '" data-bm="' + bm.id + '">' +
        '<div class="bm-preview ' + bm.id + '"></div>' +
        '<div><div class="bm-name">' + bm.name + '</div>' +
        '<div class="bm-desc">' + bm.desc + '</div></div>' +
        '</div>'
      );
    }).join('');

    menu.querySelectorAll('.bm-option').forEach(function (el) {
      el.addEventListener('click', function () {
        switchBasemap(el.dataset.bm);
        menu.classList.remove('open');
      });
    });

    /* toggle dropdown button */
    const btn = document.querySelector('.basemap-dropdown > .map-btn');
    if (btn) {
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        menu.classList.toggle('open');
      });
    }
    document.addEventListener('click', function () { menu.classList.remove('open'); });
  }

  /* ════════════════════════════════════════════
   *  CHOROPLETH LAYER
   * ════════════════════════════════════════════ */
  function computeAmphoeValues() {
    const vals = {};
    AMPHOE_LIST.forEach(function (am) { vals[am] = null; });

    if (!_summaryData) return vals;

    if (_choroDimension === 'water_pct') {
      /* เฉลี่ย % ระดับน้ำ/ตลิ่ง ของสถานีในอำเภอนั้น */
      const sums = {}, counts = {};
      (_summaryData.stations || []).forEach(function (s) {
        const am = s.amphoe;
        if (!am || !AMPHOE_LIST.includes(am)) return;
        const cur  = parseFloat(s.current_level || s.current);
        const bank = parseFloat(s.bank_level    || s.bank);
        if (isNaN(cur) || isNaN(bank) || bank <= 0) return;
        const pct = cur / bank * 100;
        sums[am]   = (sums[am]   || 0) + pct;
        counts[am] = (counts[am] || 0) + 1;
      });
      Object.keys(sums).forEach(function (am) {
        vals[am] = sums[am] / counts[am];
      });

    } else if (_choroDimension === 'rain_24') {
      /* ฝน 24 ชม. รายอำเภอ */
      (_summaryData.rain || []).forEach(function (r) {
        const am = r.amphoe;
        if (am && AMPHOE_LIST.includes(am)) {
          vals[am] = parseFloat(r.rain24 || r.rain_24hr) || 0;
        }
      });

    } else if (_choroDimension === 'risk') {
      /* risk score: นับจำนวนสถานีวิกฤติ+เฝ้าระวัง */
      const risk = {};
      (_summaryData.stations || []).forEach(function (s) {
        const am = s.amphoe;
        if (!am) return;
        risk[am] = (risk[am] || 0) + (s.status === 'วิกฤติ' ? 2 : s.status === 'เฝ้าระวัง' ? 1 : 0);
      });
      Object.keys(risk).forEach(function (am) {
        if (AMPHOE_LIST.includes(am)) vals[am] = risk[am] * 25;
      });
    }
    return vals;
  }

  function buildChoroplethLayer(amphoeBoundaries) {
    const vals = computeAmphoeValues();
    _layerChoropleth = L.geoJSON(amphoeBoundaries, {
      style: function (feature) {
        const am  = feature.properties && (feature.properties.amphoe || feature.properties.NAME_2 || '');
        const pct = vals[am];
        return {
          fillColor:   choroplethColor(pct),
          fillOpacity: 0.55,
          color:       '#fff',
          weight:      1.5,
          dashArray:   null,
        };
      },
      onEachFeature: function (feature, layer) {
        const am  = feature.properties && (feature.properties.amphoe || feature.properties.NAME_2 || '');
        const pct = vals[am];
        layer.bindTooltip(
          '<div style="font-family:Sarabun,sans-serif;font-weight:700;">' +
          'อ.' + am + '</div>' +
          '<div style="font-size:12px;color:#334155;">' +
          (pct != null ? pct.toFixed(1) + (
            _choroDimension === 'rain_24' ? ' มม.' :
            _choroDimension === 'risk'    ? ' คะแนนเสี่ยง' : '%'
          ) : 'ไม่มีข้อมูล') +
          '</div>',
          { sticky: true, className: 'nbp-river-tip' }
        );
        layer.on('mouseover', function () {
          layer.setStyle({ fillOpacity: 0.8, weight: 2.5 });
        });
        layer.on('mouseout', function () {
          _layerChoropleth && _layerChoropleth.resetStyle(layer);
        });
      }
    });
    return _layerChoropleth;
  }

  function refreshChoropleth(amphoeBoundaries) {
    if (!_layerChoropleth) return;
    const vals = computeAmphoeValues();
    _layerChoropleth.eachLayer(function (layer) {
      const f  = layer.feature;
      const am = f && f.properties && (f.properties.amphoe || f.properties.NAME_2 || '');
      const pct = vals[am];
      layer.setStyle({ fillColor: choroplethColor(pct) });
    });
  }

  /* ════════════════════════════════════════════
   *  HEATMAP LAYER
   * ════════════════════════════════════════════ */
  function buildHeatmapLayer() {
    if (!global.L || !global.L.heatLayer) return null;
    const points = [];
    if (_summaryData && _summaryData.stations) {
      _summaryData.stations.forEach(function (s) {
        const lat = parseFloat(s.lat);
        const lng = parseFloat(s.lon || s.lng);
        if (isNaN(lat) || isNaN(lng)) return;
        const cur  = parseFloat(s.current_level || s.current) || 0;
        const bank = parseFloat(s.bank_level    || s.bank)    || 1;
        const intensity = Math.min(cur / bank, 1.2); /* > 1 = ล้นตลิ่ง */
        points.push([lat, lng, intensity]);
      });
    }
    if (!points.length) return null;
    _layerHeatmap = L.heatLayer(points, {
      radius:  30,
      blur:    22,
      maxZoom: 12,
      max:     1.2,
      gradient: {
        0.0: '#3b82f6',
        0.3: '#10b981',
        0.6: '#f59e0b',
        0.85:'#ef4444',
        1.0: '#7f1d1d',
      },
    });
    return _layerHeatmap;
  }

  /* ════════════════════════════════════════════
   *  LAYER PANEL UI
   * ════════════════════════════════════════════ */
  function buildLayerPanel() {
    const toggle = document.querySelector('.lp-toggle');
    const body   = document.querySelector('.lp-body');
    if (!toggle || !body) return;

    toggle.addEventListener('click', function (e) {
      e.stopPropagation();
      const isOpen = body.classList.toggle('show');
      toggle.classList.toggle('active', isOpen);
    });
    document.addEventListener('click', function () {
      body.classList.remove('show');
      toggle.classList.remove('active');
    });
    body.addEventListener('click', function (e) { e.stopPropagation(); });

    /* checkbox handlers */
    var chkLines  = document.getElementById('chkLines');
    var chkPoints = document.getElementById('chkPoints');
    var chkFlow   = document.getElementById('chkFlow');
    var chkChoro  = document.getElementById('chkChoro');
    var chkHeat   = document.getElementById('chkHeat');
    var choroSub  = document.getElementById('choroSub');

    function syncLayer(chk, layer) {
      if (!chk || !layer || !_map) return;
      chk.addEventListener('change', function () {
        if (chk.checked) _map.addLayer(layer);
        else _map.removeLayer(layer);
      });
    }

    /* Wait for waterways:loaded event */
    document.addEventListener('nbp:waterways:loaded', function (ev) {
      _layerWaterwayLines = ev.detail.lines;
      _layerStationPoints = ev.detail.points;
      _layerFlowAnim      = ev.detail.flow;
      syncLayer(chkLines,  _layerWaterwayLines);
      syncLayer(chkPoints, _layerStationPoints);
      syncLayer(chkFlow,   _layerFlowAnim);
      updateLayerCount();
    });

    /* Choropleth toggle */
    if (chkChoro) {
      chkChoro.addEventListener('change', function () {
        if (!_map) return;
        if (chkChoro.checked) {
          if (!_layerChoropleth) {
            showToast('ต้องโหลด boundary ของอำเภอก่อน', 'warn');
            chkChoro.checked = false;
            return;
          }
          _map.addLayer(_layerChoropleth);
          if (choroSub) choroSub.style.display = 'block';
          showChoroplethLegend(true);
        } else {
          if (_layerChoropleth) _map.removeLayer(_layerChoropleth);
          if (choroSub) choroSub.style.display = 'none';
          showChoroplethLegend(false);
        }
        updateLayerCount();
      });
    }

    /* Heatmap toggle */
    if (chkHeat) {
      chkHeat.addEventListener('change', function () {
        if (!_map) return;
        if (chkHeat.checked) {
          if (!_layerHeatmap) buildHeatmapLayer();
          if (_layerHeatmap) _map.addLayer(_layerHeatmap);
        } else {
          if (_layerHeatmap) _map.removeLayer(_layerHeatmap);
        }
        updateLayerCount();
      });
    }

    /* Choropleth dimension radio */
    document.querySelectorAll('[name="choroDim"]').forEach(function (radio) {
      radio.addEventListener('change', function () {
        _choroDimension = radio.value;
        refreshChoropleth();
        updateChoroplethLegendTitle();
      });
    });
  }

  function updateLayerCount() {
    const el = document.querySelector('.lp-count');
    if (!el) return;
    var count = 0;
    document.querySelectorAll('.lp-body input[type="checkbox"]:checked').forEach(function () { count++; });
    el.textContent = count;
  }

  /* ════════════════════════════════════════════
   *  CHOROPLETH LEGEND
   * ════════════════════════════════════════════ */
  function showChoroplethLegend(show) {
    var el = document.querySelector('.choropleth-legend');
    if (!el) return;
    el.style.display = show ? 'block' : 'none';
  }

  function updateChoroplethLegendTitle() {
    var el = document.querySelector('.cl-title');
    if (!el) return;
    var titles = {
      water_pct: '% ระดับน้ำ / ตลิ่ง',
      rain_24:   'ฝน 24 ชม. (มม.)',
      risk:      'คะแนนความเสี่ยง',
    };
    el.textContent = titles[_choroDimension] || '';
  }

  /* ════════════════════════════════════════════
   *  AUTO-REFRESH COUNTDOWN
   * ════════════════════════════════════════════ */
  function initRefreshCountdown(onRefresh, periodMs) {
    _refreshPeriodMs  = periodMs || 5 * 60 * 1000;
    _refreshCountdown = _refreshPeriodMs / 1000;

    var badge = document.getElementById('refreshCountdown');

    if (_refreshInterval) clearInterval(_refreshInterval);
    _refreshInterval = setInterval(function () {
      _refreshCountdown -= 1;
      if (badge) {
        badge.textContent = _refreshCountdown > 0
          ? 'รีเฟรชใน ' + _refreshCountdown + 'วิ'
          : 'กำลังโหลด...';
      }
      if (_refreshCountdown <= 0) {
        _refreshCountdown = _refreshPeriodMs / 1000;
        if (typeof onRefresh === 'function') onRefresh();
      }
    }, 1000);
  }

  /* ════════════════════════════════════════════
   *  TOAST NOTIFICATION
   * ════════════════════════════════════════════ */
  function showToast(msg, type, durationMs) {
    type       = type       || 'info';
    durationMs = durationMs || 3000;

    var icon = { info: 'ℹ️', success: '✅', warn: '⚠️', error: '❌' }[type] || 'ℹ️';
    var el = document.createElement('div');
    el.className = 'toast' + (type === 'success' ? ' success' : type === 'error' ? ' error' : '');
    el.innerHTML = '<span style="font-size:16px;">' + icon + '</span><span>' + msg + '</span>';
    document.body.appendChild(el);
    setTimeout(function () {
      el.style.animation = 'toastIn .3s ease-out reverse forwards';
      setTimeout(function () { el.remove(); }, 300);
    }, durationMs);
  }

  /* ════════════════════════════════════════════
   *  SCALE CONTROL (เพิ่ม metric scale bar)
   * ════════════════════════════════════════════ */
  function addScaleBar() {
    if (!_map) return;
    L.control.scale({ imperial: false, position: 'bottomright' }).addTo(_map);
  }

  /* ════════════════════════════════════════════
   *  UPDATE WITH LIVE DATA
   * ════════════════════════════════════════════ */
  function updateData(summaryJSON) {
    _summaryData = summaryJSON;
    /* relay to WaterwaysLoader for marker refresh */
    if (global.WaterwaysLoader && summaryJSON && summaryJSON.stations) {
      global.WaterwaysLoader.updateStationData(summaryJSON.stations);
    }
    /* update heatmap if visible */
    if (_layerHeatmap && _map && _map.hasLayer(_layerHeatmap)) {
      _map.removeLayer(_layerHeatmap);
      buildHeatmapLayer();
      if (_layerHeatmap) _map.addLayer(_layerHeatmap);
    }
    /* update choropleth if visible */
    if (_layerChoropleth && _map && _map.hasLayer(_layerChoropleth)) {
      refreshChoropleth();
    }
  }

  /* ════════════════════════════════════════════
   *  MAIN INIT
   * ════════════════════════════════════════════ */
  /**
   * init(map, options)
   * @param {L.Map} map
   * @param {Object} opts
   *   opts.defaultBasemap    {string}   id of basemap (default: 'esri-light')
   *   opts.refreshCallback   {Function} called every refreshPeriod
   *   opts.refreshPeriodMs   {number}   ms between auto-refresh (default: 300000)
   *   opts.amphoeBoundaries  {Object}   GeoJSON FeatureCollection ของอำเภอ (optional)
   */
  function init(map, opts) {
    opts        = opts || {};
    _map        = map;
    _activeBaseId = opts.defaultBasemap || 'esri-light';

    /* Basemaps */
    initBasemaps();
    buildBasemapDropdown();

    /* Layer panel */
    buildLayerPanel();

    /* Scale bar */
    addScaleBar();

    /* Choropleth (ถ้ามี boundary ส่งมา) */
    if (opts.amphoeBoundaries) {
      buildChoroplethLayer(opts.amphoeBoundaries);
    }

    /* Heatmap (lazy — สร้างเมื่อมีข้อมูล) */

    /* Auto-refresh countdown */
    if (opts.refreshCallback) {
      initRefreshCountdown(opts.refreshCallback, opts.refreshPeriodMs);
    }

    /* Init legend hidden */
    showChoroplethLegend(false);

    return {
      switchBasemap:        switchBasemap,
      updateData:           updateData,
      buildChoroplethLayer: buildChoroplethLayer,
      buildHeatmapLayer:    buildHeatmapLayer,
      showToast:            showToast,
    };
  }

  /* ════════════════════════════════════════════
   *  EXPORTS
   * ════════════════════════════════════════════ */
  global.MapLayers = {
    init:                 init,
    switchBasemap:        switchBasemap,
    updateData:           updateData,
    buildChoroplethLayer: buildChoroplethLayer,
    buildHeatmapLayer:    buildHeatmapLayer,
    showToast:            showToast,
    initRefreshCountdown: initRefreshCountdown,
    BASEMAPS:             BASEMAPS,
  };

})(window);
