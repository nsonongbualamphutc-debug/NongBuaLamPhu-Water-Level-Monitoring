/**
 * waterways-loader.js
 * โหลด GeoJSON ลำน้ำ + จุดสถานีลงแผนที่ Leaflet
 * ระบบสถานการณ์น้ำ จ.หนองบัวลำภู — สำนักงานสถิติจังหวัดหนองบัวลำภู
 *
 * ต้องโหลดหลัง: Leaflet.js, config.js
 * ไฟล์นี้ expose: window.WaterwaysLoader
 */

(function (global) {
  'use strict';

  /* ── สีประจำแต่ละลำน้ำ ── */
  const RIVER_PALETTE = {
    paneang:  { stroke: '#0ea5e9', fill: '#bae6fd', label: 'ลำน้ำพะเนียง' },
    mong:     { stroke: '#8b5cf6', fill: '#ede9fe', label: 'ลำน้ำโมง'    },
    mo:       { stroke: '#10b981', fill: '#d1fae5', label: 'ลำน้ำมอ'     },
    phuay:    { stroke: '#f59e0b', fill: '#fef3c7', label: 'ลำน้ำพวย'    },
    default:  { stroke: '#3b82f6', fill: '#dbeafe', label: 'ลำน้ำ'       },
  };

  /* ── Status colors ── */
  const STATUS_COLOR = {
    ปกติ:       '#16a34a',
    เฝ้าระวัง:  '#f59e0b',
    วิกฤติ:     '#dc2626',
    'N/A':      '#94a3b8',
  };

  /* ── GeoJSON source paths (relative to repo root) ── */
  const BASE = (global.APP_CONFIG && global.APP_CONFIG.BASE_URL) || '';
  const GEOJSON_LINES  = BASE + 'nbp_water_lines.geojson';
  const GEOJSON_POINTS = BASE + 'nbp_water_points.geojson';

  /* ── Layer references (ใช้ MapLayers.getLayers() เพื่อ toggle ได้) ── */
  let _layerLines   = null;
  let _layerPoints  = null;
  let _layerFlow    = null;   // animated flow particles
  let _mapRef       = null;
  let _stationData  = {};     // { station_id: { current, status, ... } }

  /* ════════════════════════════════════════════
   *  UTILITY
   * ════════════════════════════════════════════ */
  function getRiverKey(name) {
    if (!name) return 'default';
    const n = String(name).toLowerCase();
    if (n.indexOf('พะเนียง') !== -1 || n.indexOf('paneang') !== -1) return 'paneang';
    if (n.indexOf('โมง')    !== -1 || n.indexOf('mong')    !== -1) return 'mong';
    if (n.indexOf('มอ')     !== -1 || n.indexOf('mo')      !== -1) return 'mo';
    if (n.indexOf('พวย')    !== -1 || n.indexOf('phuay')   !== -1) return 'phuay';
    return 'default';
  }

  function fetchJSON(url) {
    return fetch(url, { cache: 'no-cache' })
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + url);
        return r.json();
      });
  }

  /* ════════════════════════════════════════════
   *  RIVER LINE LAYER
   * ════════════════════════════════════════════ */
  function buildLineLayer(geojson) {
    _layerLines = L.geoJSON(geojson, {
      style: function (feature) {
        const key = getRiverKey(
          (feature.properties && (feature.properties.river || feature.properties.name)) || ''
        );
        const pal = RIVER_PALETTE[key] || RIVER_PALETTE.default;
        return {
          color:       pal.stroke,
          weight:      4.5,
          opacity:     0.82,
          lineCap:     'round',
          lineJoin:    'round',
          dashArray:   null,
        };
      },
      onEachFeature: function (feature, layer) {
        const p    = feature.properties || {};
        const key  = getRiverKey(p.river || p.name || '');
        const pal  = RIVER_PALETTE[key];
        const name = p.river || p.name || pal.label;
        layer.bindTooltip(
          '<div style="font-family:Sarabun,sans-serif;font-weight:700;color:#0c4a6e;">' +
          '🌊 ' + name + '</div>',
          { sticky: true, className: 'nbp-river-tip' }
        );
        /* highlight on hover */
        layer.on('mouseover', function () {
          layer.setStyle({ weight: 7, opacity: 1 });
        });
        layer.on('mouseout', function () {
          _layerLines && _layerLines.resetStyle(layer);
        });
      }
    });
    return _layerLines;
  }

  /* ════════════════════════════════════════════
   *  STATION POINT LAYER
   * ════════════════════════════════════════════ */
  function stationIcon(status, stationId) {
    const color = STATUS_COLOR[status] || STATUS_COLOR['N/A'];
    const isPulse = status === 'วิกฤติ';
    const pulseHTML = isPulse
      ? '<div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);">' +
        '<div style="width:36px;height:36px;border-radius:50%;background:' + color +
        ';opacity:.35;animation:wlPulse 1.5s ease-out infinite;"></div></div>'
      : '';

    return L.divIcon({
      className: 'nbp-station-icon',
      html:
        '<div style="position:relative;width:28px;height:28px;">' +
        pulseHTML +
        '<div style="width:28px;height:28px;border-radius:50%;background:' + color +
        ';border:3px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.25);' +
        'display:flex;align-items:center;justify-content:center;' +
        'font-size:10px;font-weight:800;color:#fff;font-family:Sarabun,sans-serif;' +
        'position:relative;z-index:1;">' +
        (stationId ? stationId.replace(/[A-Z]+/, '') : '?') +
        '</div></div>',
      iconSize:   [28, 28],
      iconAnchor: [14, 14],
      popupAnchor:[0, -16],
    });
  }

  function buildPopupHTML(props, liveData) {
    const sid    = props.station_id || props.id || '—';
    const name   = props.name || sid;
    const river  = props.river || '—';
    const amphoe = props.amphoe || '—';
    const bank   = parseFloat(props.bank_level || props.bank || 0);
    const warn   = parseFloat(props.warn_level || props.warn || 0);

    const live    = liveData || {};
    const level   = live.current != null ? parseFloat(live.current) : null;
    const status  = live.status  || (level == null ? 'N/A' : 'ปกติ');
    const color   = STATUS_COLOR[status] || '#94a3b8';
    const diff    = level != null ? (bank - level).toFixed(2) : '—';
    const lvTxt   = level != null ? level.toFixed(2) + ' ม.รทก.' : 'ไม่มีข้อมูล';
    const pct     = (level != null && bank > 0) ? Math.min(Math.round(level / bank * 100), 100) : 0;

    const riverKey = getRiverKey(river);
    const page     = { paneang: 'paneang', mong: 'mong', mo: 'mo', phuay: 'phuay' }[riverKey] || 'index';

    return (
      '<div style="font-family:Sarabun,sans-serif;min-width:220px;">' +
      '<div style="background:linear-gradient(135deg,#0c4a6e,#0284c7);color:#fff;' +
      'padding:10px 14px;margin:-10px -14px 12px;border-radius:8px 8px 0 0;">' +
      '<div style="font-weight:800;font-size:14px;">🏔 ' + name + '</div>' +
      '<div style="font-size:11px;opacity:.8;margin-top:2px;">🌊 ' + river + ' · อ.' + amphoe + '</div>' +
      '</div>' +
      /* Level */
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">' +
      '<div>' +
      '<div style="font-size:11px;color:#64748b;">ระดับน้ำปัจจุบัน</div>' +
      '<div style="font-size:22px;font-weight:800;color:' + color + ';">' + lvTxt + '</div>' +
      '</div>' +
      '<div style="background:' + color + '22;border:2px solid ' + color + ';' +
      'border-radius:8px;padding:6px 10px;text-align:center;">' +
      '<div style="font-size:13px;font-weight:800;color:' + color + ';">' + status + '</div>' +
      '</div></div>' +
      /* Progress bar */
      (level != null ? (
        '<div style="background:#f1f5f9;border-radius:4px;height:8px;margin-bottom:10px;overflow:hidden;">' +
        '<div style="width:' + pct + '%;height:100%;background:' + color + ';border-radius:4px;transition:width .6s;"></div></div>'
      ) : '') +
      /* Info grid */
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px;">' +
      '<div style="background:#f8fafc;border-radius:8px;padding:8px;">' +
      '<div style="font-size:10px;color:#64748b;">ตลิ่ง (ม.รทก.)</div>' +
      '<div style="font-weight:700;color:#dc2626;">' + (bank > 0 ? bank.toFixed(2) : '—') + '</div></div>' +
      '<div style="background:#f8fafc;border-radius:8px;padding:8px;">' +
      '<div style="font-size:10px;color:#64748b;">เฝ้าระวัง</div>' +
      '<div style="font-weight:700;color:#f59e0b;">' + (warn > 0 ? warn.toFixed(2) : '—') + '</div></div>' +
      '<div style="background:#f8fafc;border-radius:8px;padding:8px;">' +
      '<div style="font-size:10px;color:#64748b;">ต่ำกว่าตลิ่ง</div>' +
      '<div style="font-weight:700;color:#0284c7;">' + diff + ' ม.</div></div>' +
      '<div style="background:#f8fafc;border-radius:8px;padding:8px;">' +
      '<div style="font-size:10px;color:#64748b;">รหัสสถานี</div>' +
      '<div style="font-weight:700;font-family:monospace;">' + sid + '</div></div>' +
      '</div>' +
      /* CTA */
      '<a href="' + page + '.html" style="display:block;text-align:center;' +
      'background:linear-gradient(135deg,#0284c7,#0ea5e9);color:#fff;' +
      'padding:8px;border-radius:8px;font-weight:700;font-size:12px;text-decoration:none;">' +
      '🌊 ดูหน้าลำน้ำ</a>' +
      '</div>'
    );
  }

  function buildPointLayer(geojson) {
    if (!global.L || !global.L.markerClusterGroup) {
      /* fallback to plain layer if MarkerCluster not loaded */
      return buildPointLayerPlain(geojson);
    }
    const cluster = L.markerClusterGroup({
      maxClusterRadius: 45,
      showCoverageOnHover: false,
      iconCreateFunction: function (c) {
        const n = c.getChildCount();
        return L.divIcon({
          html: '<div style="width:38px;height:38px;border-radius:50%;' +
            'background:linear-gradient(135deg,#0284c7,#0ea5e9);color:#fff;' +
            'display:flex;align-items:center;justify-content:center;' +
            'font-weight:800;font-size:13px;border:3px solid #fff;' +
            'box-shadow:0 2px 10px rgba(0,0,0,.25);font-family:Sarabun,sans-serif;">' +
            n + '</div>',
          className: 'nbp-cluster',
          iconSize: [38, 38],
          iconAnchor: [19, 19],
        });
      }
    });

    L.geoJSON(geojson, {
      pointToLayer: function (feature, latlng) {
        const p    = feature.properties || {};
        const sid  = p.station_id || p.id || '';
        const live = _stationData[sid] || {};
        const marker = L.marker(latlng, {
          icon: stationIcon(live.status || 'N/A', sid),
          riseOnHover: true,
        });
        marker.bindPopup(buildPopupHTML(p, live), {
          maxWidth: 280,
          className: 'nbp-popup',
        });
        /* update marker when data refreshes */
        marker._nbpStationId = sid;
        marker._nbpFeature   = feature;
        cluster.addLayer(marker);
        return marker;
      }
    });

    _layerPoints = cluster;
    return cluster;
  }

  function buildPointLayerPlain(geojson) {
    _layerPoints = L.geoJSON(geojson, {
      pointToLayer: function (feature, latlng) {
        const p    = feature.properties || {};
        const sid  = p.station_id || p.id || '';
        const live = _stationData[sid] || {};
        return L.marker(latlng, {
          icon: stationIcon(live.status || 'N/A', sid),
          riseOnHover: true,
        }).bindPopup(buildPopupHTML(p, live), { maxWidth: 280, className: 'nbp-popup' });
      }
    });
    return _layerPoints;
  }

  /* ════════════════════════════════════════════
   *  FLOW ANIMATION (เส้นน้ำไหล ด้วย Canvas marker)
   * ════════════════════════════════════════════ */
  function buildFlowLayer(geojson) {
    /* สร้าง animated dots วิ่งตามเส้นแม่น้ำ */
    const dots = [];
    if (!geojson || !geojson.features) return null;

    geojson.features.forEach(function (f) {
      if (!f.geometry) return;
      const coords = f.geometry.type === 'LineString'
        ? f.geometry.coordinates
        : (f.geometry.type === 'MultiLineString' ? f.geometry.coordinates[0] : null);
      if (!coords || coords.length < 2) return;

      const key = getRiverKey(
        (f.properties && (f.properties.river || f.properties.name)) || ''
      );
      const color = (RIVER_PALETTE[key] || RIVER_PALETTE.default).stroke;

      /* เลือกจุดกลางๆ ของเส้น */
      const steps = [0.2, 0.4, 0.6, 0.8];
      steps.forEach(function (t) {
        const idx = Math.floor(t * (coords.length - 1));
        const c   = coords[idx];
        if (!c) return;
        dots.push({
          lat:    c[1],
          lng:    c[0],
          color:  color,
          phase:  Math.random() * Math.PI * 2,
          speed:  0.4 + Math.random() * 0.6,
        });
      });
    });

    /* ใช้ Canvas renderer วาด circle markers แอนิเมต */
    const renderer = L.canvas({ padding: 0.5 });
    const circles  = dots.map(function (d) {
      return L.circleMarker([d.lat, d.lng], {
        renderer:    renderer,
        radius:      4,
        color:       d.color,
        fillColor:   d.color,
        fillOpacity: 0.75,
        weight:      0,
        interactive: false,
      });
    });

    _layerFlow = L.layerGroup(circles);

    /* pulse animation via requestAnimationFrame */
    let _raf;
    function animate(ts) {
      if (!_mapRef) return;
      circles.forEach(function (c, i) {
        const d   = dots[i];
        const osc = 0.4 + 0.4 * Math.sin(ts * 0.001 * d.speed + d.phase);
        c.setStyle({ fillOpacity: osc, radius: 2 + osc * 3 });
      });
      _raf = requestAnimationFrame(animate);
    }
    _layerFlow.on('add',    function () { _raf = requestAnimationFrame(animate); });
    _layerFlow.on('remove', function () { cancelAnimationFrame(_raf); });

    return _layerFlow;
  }

  /* ════════════════════════════════════════════
   *  INJECT POPUP / TOOLTIP CSS
   * ════════════════════════════════════════════ */
  function injectStyles() {
    if (document.getElementById('nbp-waterways-style')) return;
    const el = document.createElement('style');
    el.id = 'nbp-waterways-style';
    el.textContent = [
      '.nbp-river-tip { background:#fff; border:1px solid #bae6fd; border-radius:8px;',
      '  padding:5px 10px; box-shadow:0 4px 12px rgba(0,0,0,.12); }',
      '.nbp-popup .leaflet-popup-content-wrapper {',
      '  border-radius:12px; padding:10px 14px; box-shadow:0 8px 28px rgba(0,0,0,.18); }',
      '.nbp-popup .leaflet-popup-tip-container { margin-top:-1px; }',
      '.nbp-station-icon { background:transparent !important; border:none !important; }',
      '.nbp-cluster { background:transparent !important; border:none !important; }',
      '@keyframes wlPulse {',
      '  0%  { transform:translate(-50%,-50%) scale(1); opacity:.35; }',
      '  100%{ transform:translate(-50%,-50%) scale(2.4); opacity:0; }',
      '}',
    ].join('\n');
    document.head.appendChild(el);
  }

  /* ════════════════════════════════════════════
   *  UPDATE LIVE DATA (เรียกจาก dashboard หลังดึง API)
   * ════════════════════════════════════════════ */
  function updateStationData(summaryStations) {
    if (!summaryStations) return;
    summaryStations.forEach(function (s) {
      const sid = s.station_id || s.id;
      if (sid) _stationData[sid] = s;
    });
    /* refresh marker icons + popups */
    if (_layerPoints) {
      if (_layerPoints.eachLayer) {
        _layerPoints.eachLayer(function (cluster) {
          if (cluster.eachLayer) {
            /* MarkerClusterGroup */
            cluster.eachLayer(function (marker) {
              if (!marker._nbpStationId) return;
              const sid  = marker._nbpStationId;
              const live = _stationData[sid] || {};
              const p    = (marker._nbpFeature && marker._nbpFeature.properties) || {};
              marker.setIcon(stationIcon(live.status || 'N/A', sid));
              marker.setPopupContent(buildPopupHTML(p, live));
            });
          } else {
            /* plain geoJSON */
            if (!cluster._nbpStationId) return;
            const sid  = cluster._nbpStationId;
            const live = _stationData[sid] || {};
            const p    = (cluster.feature && cluster.feature.properties) || {};
            cluster.setIcon(stationIcon(live.status || 'N/A', sid));
            cluster.setPopupContent(buildPopupHTML(p, live));
          }
        });
      }
    }
  }

  /* ════════════════════════════════════════════
   *  PUBLIC INIT
   * ════════════════════════════════════════════ */
  /**
   * init(map, options)
   * @param {L.Map} map   - Leaflet map instance
   * @param {Object} opts - { addToMap:bool, flowLayer:bool }
   * @returns Promise<{ lines, points, flow }>
   */
  function init(map, opts) {
    opts      = opts || {};
    _mapRef   = map;
    const addToMap  = opts.addToMap  !== false;
    const showFlow  = opts.flowLayer !== false;

    injectStyles();

    return Promise.all([fetchJSON(GEOJSON_LINES), fetchJSON(GEOJSON_POINTS)])
      .then(function (results) {
        const geojsonLines  = results[0];
        const geojsonPoints = results[1];

        _layerLines  = buildLineLayer(geojsonLines);
        _layerPoints = buildPointLayer(geojsonPoints);

        if (showFlow) {
          _layerFlow = buildFlowLayer(geojsonLines);
        }

        if (addToMap && map) {
          _layerLines.addTo(map);
          _layerPoints.addTo(map);
          if (_layerFlow) _layerFlow.addTo(map);
        }

        /* emit custom event so MapLayers can register these */
        const ev = new CustomEvent('nbp:waterways:loaded', {
          detail: {
            lines:  _layerLines,
            points: _layerPoints,
            flow:   _layerFlow,
          }
        });
        document.dispatchEvent(ev);

        return { lines: _layerLines, points: _layerPoints, flow: _layerFlow };
      })
      .catch(function (err) {
        console.error('[WaterwaysLoader] ไม่สามารถโหลด GeoJSON:', err);
        /* emit error event */
        document.dispatchEvent(new CustomEvent('nbp:waterways:error', { detail: err }));
        return { lines: null, points: null, flow: null };
      });
  }

  /* ════════════════════════════════════════════
   *  EXPORTS
   * ════════════════════════════════════════════ */
  global.WaterwaysLoader = {
    init:              init,
    updateStationData: updateStationData,
    getLayers: function () {
      return { lines: _layerLines, points: _layerPoints, flow: _layerFlow };
    },
    RIVER_PALETTE:  RIVER_PALETTE,
    STATUS_COLOR:   STATUS_COLOR,
  };

})(window);
