/**
 * Generates the in-app PDF viewer HTML used by the WebView.
 * Annotation coordinates are stored as normalized ratios so they can be
 * persisted, re-rendered across screen sizes, and burned into a saved PDF.
 */
export function getPdfViewerHtml(base64Data, options = {}) {
  const initialAnnotations = JSON.stringify(options.initialAnnotations || {});
  const initialMode = JSON.stringify(options.initialMode || 'view');
  const initialDrawColor = JSON.stringify(options.initialDrawColor || '#FF0000');
  const initialHighlightColor = JSON.stringify(
    options.initialHighlightColor || '#FF000066'
  );
  const viewerSlot = Math.max(1, Math.min(2, Number(options.viewerSlot) || 1));
  const crossPaneDragEnabled = !!options.crossPaneDragEnabled;
  const dualLayout =
    options.dualLayout === 'column' ? 'column' : 'row';
  const initialSelectedPage = Math.max(1, Number(options.initialSelectedPage) || 1);
  const initialDrawBrushPx = Math.max(
    1.5,
    Math.min(28, Number(options.initialDrawBrushPx) || 5)
  );
  const initialMarkerBrushPx = Math.max(
    6,
    Math.min(40, Number(options.initialMarkerBrushPx) || 16)
  );

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta
    name="viewport"
    content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes"
  />
  <title>PDF Viewer</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      background: #525659;
      overflow-x: hidden;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif;
      -webkit-tap-highlight-color: transparent;
    }
    #pdf-container {
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 12px 0 20px;
      gap: 8px;
    }
    .page-wrapper {
      position: relative;
      background: white;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
      margin: 0 auto;
      overflow: hidden;
      transition: box-shadow 140ms ease, transform 140ms ease;
    }
    .page-wrapper.selected {
      box-shadow:
        0 0 0 3px rgba(255, 255, 255, 0.98),
        0 0 0 6px rgba(25, 118, 210, 0.88),
        0 14px 30px rgba(0, 0, 0, 0.35);
      transform: translateY(-1px);
    }
    .page-wrapper.page-wrapper-lift {
      z-index: 3;
      transform: translateY(-8px) scale(1.02);
      box-shadow:
        0 0 0 3px rgba(100, 181, 246, 0.95),
        0 18px 36px rgba(0, 0, 0, 0.45);
      transition: transform 160ms ease, box-shadow 160ms ease;
    }
    .page-wrapper canvas {
      display: block;
      width: 100%;
      height: 100%;
    }
    .annotation-layer {
      position: absolute;
      inset: 0;
      pointer-events: none;
    }
    .annotation-layer.active {
      pointer-events: auto;
    }
    .annotation-layer.editing {
      cursor: crosshair;
      touch-action: none;
    }
    .highlight-rect {
      position: absolute;
      border-radius: 4px;
      box-shadow: 0 0 0 1px rgba(255, 255, 255, 0.24);
      pointer-events: auto;
      transition: box-shadow 120ms ease, transform 120ms ease;
    }
    .text-note {
      position: absolute;
      background: #fff9c4;
      border: 1px solid #f9a825;
      border-radius: 8px;
      padding: 6px 8px;
      font-size: 12px;
      line-height: 1.35;
      color: #333;
      max-width: 210px;
      word-wrap: break-word;
      pointer-events: auto;
      box-shadow: 0 3px 10px rgba(0, 0, 0, 0.18);
      transition: box-shadow 120ms ease, transform 120ms ease;
    }
    .highlight-rect.selected,
    .text-note.selected {
      box-shadow:
        0 0 0 2px rgba(255, 255, 255, 0.95),
        0 0 0 4px rgba(25, 118, 210, 0.92),
        0 10px 22px rgba(0, 0, 0, 0.24);
      transform: translateY(-1px);
      cursor: grab;
    }
    .highlight-rect.dragging,
    .text-note.dragging {
      cursor: grabbing;
    }
    .draw-canvas {
      position: absolute;
      inset: 0;
      pointer-events: none;
    }
    .draw-canvas.active {
      pointer-events: auto;
      cursor: crosshair;
      touch-action: none;
    }
    .page-label {
      text-align: center;
      color: #aaa;
      font-size: 11px;
      padding: 2px 0 6px;
      transition: color 140ms ease;
    }
    .page-label-row {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      flex-wrap: wrap;
      padding: 2px 8px 8px;
      width: 100%;
      max-width: 100%;
    }
    .page-label-text {
      color: #aaa;
      font-size: 11px;
      transition: color 140ms ease;
    }
    .page-label-row.selected .page-label-text {
      color: #f3f7ff;
    }
    .page-drag-handle {
      flex-shrink: 0;
      touch-action: none;
      padding: 6px 10px;
      min-width: 40px;
      background: rgba(255, 255, 255, 0.1);
      border-radius: 8px;
      border: 1px solid rgba(255, 255, 255, 0.2);
      color: #c5cae9;
      font-size: 16px;
      line-height: 1;
      cursor: grab;
    }
    .page-drag-handle:active {
      background: rgba(25, 118, 210, 0.45);
      border-color: rgba(100, 181, 246, 0.9);
    }
    .page-menu-btn {
      flex-shrink: 0;
      touch-action: manipulation;
      padding: 6px 10px;
      min-width: 36px;
      background: rgba(255, 255, 255, 0.1);
      border-radius: 8px;
      border: 1px solid rgba(255, 255, 255, 0.2);
      color: #e0e0e0;
      font-size: 18px;
      line-height: 1;
      cursor: pointer;
    }
    .page-menu-btn:active {
      background: rgba(25, 118, 210, 0.4);
    }
    .page-menu-backdrop {
      position: fixed;
      inset: 0;
      z-index: 99998;
      background: transparent;
    }
    .page-menu-dropdown {
      position: fixed;
      z-index: 99999;
      min-width: 200px;
      background: #fafafa;
      border-radius: 10px;
      box-shadow: 0 8px 28px rgba(0, 0, 0, 0.35);
      overflow: hidden;
      border: 1px solid rgba(0, 0, 0, 0.12);
    }
    .page-menu-item {
      padding: 14px 16px;
      font-size: 14px;
      color: #222;
      border: none;
      width: 100%;
      text-align: left;
      background: #fff;
      cursor: pointer;
    }
    .page-menu-item:active {
      background: #e3f2fd;
    }
    .page-menu-item.danger {
      color: #b71c1c;
    }
    .page-wrapper.page-wrapper-dimmed {
      opacity: 0.38;
      transition: opacity 140ms ease;
    }
    .page-drag-ghost {
      pointer-events: none;
    }
    .page-label.selected {
      color: #f3f7ff;
    }
    #loading {
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      color: white;
      font-size: 16px;
    }
    #error-msg {
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      color: #ff5252;
      font-size: 14px;
      text-align: center;
      padding: 20px;
      display: none;
      max-width: 340px;
    }
  </style>
</head>
<body>
  <div id="loading">Loading PDF...</div>
  <div id="error-msg"></div>
  <div id="pdf-container"></div>

  <script>
    let currentMode = ${initialMode};
    let drawColor = ${initialDrawColor};
    let drawWidth = ${initialDrawBrushPx};
    let markerBrushPx = ${initialMarkerBrushPx};
    let highlightColor = ${initialHighlightColor};
    let annotations = normalizeAnnotations(${initialAnnotations});
    let pdfDoc = null;
    let totalPages = 0;
    let selectedPage = ${initialSelectedPage};
    let selectedAnnotation = null;
    let activeDrag = null;
    let suppressNextClick = false;
    const pageMetrics = {};
    const VIEWER_SLOT = ${viewerSlot};
    const CROSS_PANE_DRAG = ${crossPaneDragEnabled ? 'true' : 'false'};
    const DUAL_LAYOUT = '${dualLayout}';
    let crossPdfLastTap = { time: 0, page: 0 };
    let pageGhostDrag = null;
    let pageMenuOpenFor = 0;

    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    function clamp(value, min, max) {
      return Math.min(max, Math.max(min, value));
    }

    function safeNumber(value, fallback) {
      const numeric = Number(value);
      return Number.isFinite(numeric) ? numeric : fallback;
    }

    function createAnnotationId(prefix) {
      return prefix + '-' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
    }

    function createEmptyPageAnnotations() {
      return { highlights: [], drawings: [], notes: [] };
    }

    function normalizePoint(point) {
      return {
        x: clamp(safeNumber(point && point.x, 0), 0, 1),
        y: clamp(safeNumber(point && point.y, 0), 0, 1),
      };
    }

    function normalizeAnnotations(rawAnnotations) {
      const normalized = {};
      if (!rawAnnotations || typeof rawAnnotations !== 'object') {
        return normalized;
      }

      Object.entries(rawAnnotations).forEach(([pageKey, pageAnnotations]) => {
        if (!pageAnnotations || typeof pageAnnotations !== 'object') {
          normalized[pageKey] = createEmptyPageAnnotations();
          return;
        }

        normalized[pageKey] = {
          highlights: Array.isArray(pageAnnotations.highlights)
            ? pageAnnotations.highlights.map((annotation) => ({
                id: annotation.id || createAnnotationId('highlight'),
                x: clamp(safeNumber(annotation.x, 0), 0, 1),
                y: clamp(safeNumber(annotation.y, 0), 0, 1),
                width: clamp(safeNumber(annotation.width, 0), 0, 1),
                height: clamp(safeNumber(annotation.height, 0), 0, 1),
                color: annotation.color || highlightColor,
              }))
            : [],
          drawings: Array.isArray(pageAnnotations.drawings)
            ? pageAnnotations.drawings
                .map((annotation) => ({
                  id: annotation.id || createAnnotationId('drawing'),
                  kind: annotation.kind === 'marker' ? 'marker' : 'pen',
                  color: annotation.color || drawColor,
                  width: clamp(safeNumber(annotation.width, 0.006), 0.001, 0.09),
                  points: Array.isArray(annotation.points)
                    ? annotation.points.map(normalizePoint)
                    : [],
                }))
                .filter((annotation) => annotation.points.length > 0)
            : [],
          notes: Array.isArray(pageAnnotations.notes)
            ? pageAnnotations.notes.map((annotation) => ({
                id: annotation.id || createAnnotationId('note'),
                x: clamp(safeNumber(annotation.x, 0), 0, 1),
                y: clamp(safeNumber(annotation.y, 0), 0, 1),
                text: String(annotation.text || ''),
              }))
            : [],
        };
      });

      return normalized;
    }

    function ensurePageAnnotations(pageNum) {
      if (!annotations[pageNum]) {
        annotations[pageNum] = createEmptyPageAnnotations();
      }
      return annotations[pageNum];
    }

    function getPageMetric(pageNum) {
      return pageMetrics[pageNum];
    }

    function getDisplayPoint(point, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric) {
        return { x: 0, y: 0 };
      }
      return {
        x: point.x * metric.displayWidth,
        y: point.y * metric.displayHeight,
      };
    }

    function getCanvasPoint(point, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric) {
        return { x: 0, y: 0 };
      }
      return {
        x: point.x * metric.drawCanvasWidth,
        y: point.y * metric.drawCanvasHeight,
      };
    }

    function pointToSegmentDistance(point, start, end) {
      const dx = end.x - start.x;
      const dy = end.y - start.y;
      if (dx === 0 && dy === 0) {
        const deltaX = point.x - start.x;
        const deltaY = point.y - start.y;
        return Math.sqrt(deltaX * deltaX + deltaY * deltaY);
      }

      const t = clamp(
        ((point.x - start.x) * dx + (point.y - start.y) * dy) / (dx * dx + dy * dy),
        0,
        1
      );

      const projection = {
        x: start.x + t * dx,
        y: start.y + t * dy,
      };

      const deltaX = point.x - projection.x;
      const deltaY = point.y - projection.y;
      return Math.sqrt(deltaX * deltaX + deltaY * deltaY);
    }

    function getStrokeBounds(points, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric || !points.length) {
        return null;
      }

      let minX = Infinity;
      let minY = Infinity;
      let maxX = -Infinity;
      let maxY = -Infinity;

      points.forEach((point) => {
        const canvasPoint = getCanvasPoint(point, pageNum);
        minX = Math.min(minX, canvasPoint.x);
        minY = Math.min(minY, canvasPoint.y);
        maxX = Math.max(maxX, canvasPoint.x);
        maxY = Math.max(maxY, canvasPoint.y);
      });

      return {
        x: minX,
        y: minY,
        width: maxX - minX,
        height: maxY - minY,
      };
    }

    function getNormalizedBounds(points) {
      if (!points.length) {
        return null;
      }

      let minX = Infinity;
      let minY = Infinity;
      let maxX = -Infinity;
      let maxY = -Infinity;

      points.forEach((point) => {
        minX = Math.min(minX, point.x);
        minY = Math.min(minY, point.y);
        maxX = Math.max(maxX, point.x);
        maxY = Math.max(maxY, point.y);
      });

      return {
        minX,
        minY,
        maxX,
        maxY,
      };
    }

    function isSelectedAnnotation(type, pageNum, id) {
      return (
        !!selectedAnnotation &&
        selectedAnnotation.type === type &&
        selectedAnnotation.page === Number(pageNum) &&
        selectedAnnotation.id === id
      );
    }

    function getAnnotationBucket(pageAnnotations, type) {
      if (type === 'highlight') {
        return pageAnnotations.highlights;
      }
      if (type === 'note') {
        return pageAnnotations.notes;
      }
      if (type === 'drawing') {
        return pageAnnotations.drawings;
      }
      return [];
    }

    function updateElementPositionFromAnnotation(element, annotation, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric || !element) {
        return;
      }

      element.style.left = annotation.x * metric.displayWidth + 'px';
      element.style.top = annotation.y * metric.displayHeight + 'px';
    }

    function beginDrag(config, event) {
      activeDrag = {
        ...config,
        hasMoved: false,
        startCoords: getEventCoords(event),
      };

      if (config.element) {
        config.element.classList.add('dragging');
      }
    }

    function updateActiveDrag(event) {
      if (!activeDrag) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();

      const currentCoords = getEventCoords(event);
      const deltaClientX = currentCoords.clientX - activeDrag.startCoords.clientX;
      const deltaClientY = currentCoords.clientY - activeDrag.startCoords.clientY;

      if (
        !activeDrag.hasMoved &&
        Math.sqrt(deltaClientX * deltaClientX + deltaClientY * deltaClientY) < 3
      ) {
        return;
      }

      activeDrag.hasMoved = true;
      const metric = getPageMetric(activeDrag.pageNum);
      if (!metric) {
        return;
      }

      const deltaX = deltaClientX / metric.displayWidth;
      const deltaY = deltaClientY / metric.displayHeight;

      if (activeDrag.kind === 'element') {
        activeDrag.annotation.x = clamp(
          activeDrag.startX + deltaX,
          0,
          Math.max(1 - activeDrag.width, 0)
        );
        activeDrag.annotation.y = clamp(
          activeDrag.startY + deltaY,
          0,
          Math.max(1 - activeDrag.height, 0)
        );

        updateElementPositionFromAnnotation(
          activeDrag.element,
          activeDrag.annotation,
          activeDrag.pageNum
        );
        return;
      }

      if (activeDrag.kind === 'drawing') {
        const bounds = getNormalizedBounds(activeDrag.originalPoints);
        if (!bounds) {
          return;
        }

        const clampedDeltaX = clamp(deltaX, -bounds.minX, 1 - bounds.maxX);
        const clampedDeltaY = clamp(deltaY, -bounds.minY, 1 - bounds.maxY);

        activeDrag.annotation.points = activeDrag.originalPoints.map((point) => ({
          x: clamp(point.x + clampedDeltaX, 0, 1),
          y: clamp(point.y + clampedDeltaY, 0, 1),
        }));

        redrawDrawings(activeDrag.pageNum);
      }
    }

    function endActiveDrag(event) {
      if (!activeDrag) {
        return;
      }

      if (event) {
        event.preventDefault();
        event.stopPropagation();
      }

      const shouldPersist = activeDrag.hasMoved;

      if (activeDrag.element) {
        activeDrag.element.classList.remove('dragging');
      }

      const pageNum = activeDrag.pageNum;
      activeDrag = null;

      if (shouldPersist) {
        suppressNextClick = true;
        if (pageNum) {
          renderPageAnnotations(pageNum);
        }
        persistAnnotations();
      }
    }

    function findDrawingAtPoint(pageNum, normalizedPoint) {
      const pageAnnotations = ensurePageAnnotations(pageNum);
      const metric = getPageMetric(pageNum);
      if (!metric) {
        return null;
      }

      const hitPoint = getDisplayPoint(normalizedPoint, pageNum);

      for (let index = pageAnnotations.drawings.length - 1; index >= 0; index -= 1) {
        const stroke = pageAnnotations.drawings[index];
        const displayLineWidth = Math.max(stroke.width * metric.displayWidth, 10);
        const threshold = displayLineWidth * 0.9 + 8;
        const points = stroke.points.map((point) => getDisplayPoint(point, pageNum));

        if (points.length === 1) {
          const deltaX = points[0].x - hitPoint.x;
          const deltaY = points[0].y - hitPoint.y;
          if (Math.sqrt(deltaX * deltaX + deltaY * deltaY) <= threshold) {
            return stroke;
          }
          continue;
        }

        for (let pointIndex = 0; pointIndex < points.length - 1; pointIndex += 1) {
          if (
            pointToSegmentDistance(hitPoint, points[pointIndex], points[pointIndex + 1]) <=
            threshold
          ) {
            return stroke;
          }
        }
      }

      return null;
    }

    function updateDomSelection() {
      updatePageSelection();

      document.querySelectorAll('.highlight-rect, .text-note').forEach((element) => {
        const isSelected =
          !!selectedAnnotation &&
          element.dataset.annotationId === selectedAnnotation.id &&
          element.dataset.annotationType === selectedAnnotation.type &&
          element.dataset.page === String(selectedAnnotation.page);
        element.classList.toggle('selected', isSelected);
      });

      for (let pageNum = 1; pageNum <= totalPages; pageNum += 1) {
        redrawDrawings(pageNum);
      }
    }

    function updatePageSelection() {
      document.querySelectorAll('.page-wrapper').forEach((element) => {
        element.classList.toggle('selected', Number(element.dataset.page) === selectedPage);
      });

      document.querySelectorAll('.page-label').forEach((element) => {
        element.classList.toggle('selected', Number(element.dataset.page) === selectedPage);
      });
    }

    function setSelectedPage(pageNum, silent) {
      if (!totalPages) {
        return;
      }

      const nextPage = clamp(safeNumber(pageNum, 1), 1, totalPages);
      const didChange = nextPage !== selectedPage;
      selectedPage = nextPage;
      updatePageSelection();

      if (!silent && didChange) {
        sendMessage({
          type: 'pageSelected',
          page: selectedPage,
        });
      }
    }

    function selectAnnotation(annotation) {
      if (annotation) {
        setSelectedPage(annotation.page);
        selectedAnnotation = {
          id: annotation.id,
          type: annotation.type,
          page: Number(annotation.page),
        };
      } else {
        selectedAnnotation = null;
      }

      updateDomSelection();
      sendMessage({
        type: 'annotationSelected',
        annotation: selectedAnnotation,
      });
    }

    function persistAnnotations() {
      sendMessage({
        type: 'annotationsChanged',
        annotations,
      });
    }

    function createHighlightElement(annotation, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric) {
        return null;
      }

      const highlight = document.createElement('div');
      highlight.className = 'highlight-rect';
      highlight.dataset.annotationId = annotation.id;
      highlight.dataset.annotationType = 'highlight';
      highlight.dataset.page = pageNum;
      highlight.style.left = annotation.x * metric.displayWidth + 'px';
      highlight.style.top = annotation.y * metric.displayHeight + 'px';
      highlight.style.width = annotation.width * metric.displayWidth + 'px';
      highlight.style.height = annotation.height * metric.displayHeight + 'px';
      highlight.style.background = annotation.color;

      function handleDragStart(event) {
        if (currentMode !== 'view') {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        if (!isSelectedAnnotation('highlight', pageNum, annotation.id)) {
          selectAnnotation({
            id: annotation.id,
            type: 'highlight',
            page: pageNum,
          });
          return;
        }

        beginDrag(
          {
            kind: 'element',
            annotation,
            element: highlight,
            pageNum,
            startX: annotation.x,
            startY: annotation.y,
            width: annotation.width,
            height: annotation.height,
          },
          event
        );
      }

      highlight.addEventListener('touchstart', handleDragStart, { passive: false });
      highlight.addEventListener('mousedown', handleDragStart);

      highlight.addEventListener('click', (event) => {
        if (currentMode !== 'view') {
          return;
        }
        if (suppressNextClick) {
          suppressNextClick = false;
          event.preventDefault();
          event.stopPropagation();
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        selectAnnotation({
          id: annotation.id,
          type: 'highlight',
          page: pageNum,
        });
      });

      return highlight;
    }

    function createNoteElement(annotation, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric) {
        return null;
      }

      const note = document.createElement('div');
      note.className = 'text-note';
      note.dataset.annotationId = annotation.id;
      note.dataset.annotationType = 'note';
      note.dataset.page = pageNum;
      note.textContent = annotation.text;

      const maxX = Math.max(metric.displayWidth - 20, 0);
      const maxY = Math.max(metric.displayHeight - 20, 0);
      note.style.left = Math.min(annotation.x * metric.displayWidth, maxX) + 'px';
      note.style.top = Math.min(annotation.y * metric.displayHeight, maxY) + 'px';

      function handleDragStart(event) {
        if (currentMode !== 'view') {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        if (!isSelectedAnnotation('note', pageNum, annotation.id)) {
          selectAnnotation({
            id: annotation.id,
            type: 'note',
            page: pageNum,
          });
          return;
        }

        beginDrag(
          {
            kind: 'element',
            annotation,
            element: note,
            pageNum,
            startX: annotation.x,
            startY: annotation.y,
            width: Math.min(note.offsetWidth / metric.displayWidth, 1),
            height: Math.min(note.offsetHeight / metric.displayHeight, 1),
          },
          event
        );
      }

      note.addEventListener('touchstart', handleDragStart, { passive: false });
      note.addEventListener('mousedown', handleDragStart);

      note.addEventListener('click', (event) => {
        if (currentMode !== 'view') {
          return;
        }
        if (suppressNextClick) {
          suppressNextClick = false;
          event.preventDefault();
          event.stopPropagation();
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        selectAnnotation({
          id: annotation.id,
          type: 'note',
          page: pageNum,
        });
      });

      return note;
    }

    function renderPageAnnotations(pageNum) {
      const overlay = document.querySelector('.annotation-layer[data-page="' + pageNum + '"]');
      if (!overlay) {
        return;
      }

      overlay.querySelectorAll('.highlight-rect, .text-note').forEach((element) => element.remove());

      const pageAnnotations = ensurePageAnnotations(pageNum);
      pageAnnotations.highlights.forEach((annotation) => {
        const highlight = createHighlightElement(annotation, pageNum);
        if (highlight) {
          overlay.appendChild(highlight);
        }
      });

      pageAnnotations.notes.forEach((annotation) => {
        const note = createNoteElement(annotation, pageNum);
        if (note) {
          overlay.appendChild(note);
        }
      });

      redrawDrawings(pageNum);
      updateDomSelection();
    }

    function drawStrokeToCanvas(ctx, stroke, pageNum, options) {
      const metric = getPageMetric(pageNum);
      if (!metric || !stroke.points.length) {
        return;
      }

      const points = stroke.points.map((point) => getCanvasPoint(point, pageNum));
      const baseLineWidth = Math.max(stroke.width * metric.drawCanvasWidth, 1.5);
      const isMarker = stroke.kind === 'marker';
      const isSelected = !!(options && options.selected);

      ctx.save();
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      if (isMarker) {
        ctx.globalCompositeOperation = 'multiply';
        ctx.globalAlpha = 0.5;
      }
      ctx.strokeStyle = stroke.color;
      ctx.lineWidth = baseLineWidth;

      if (isSelected && !isMarker) {
        ctx.strokeStyle = 'rgba(255, 255, 255, 0.95)';
        ctx.lineWidth = baseLineWidth + 8;
        strokePath(ctx, points);
        ctx.stroke();

        ctx.strokeStyle = 'rgba(25, 118, 210, 0.92)';
        ctx.lineWidth = baseLineWidth + 4;
        strokePath(ctx, points);
        ctx.stroke();

        ctx.strokeStyle = stroke.color;
        ctx.lineWidth = baseLineWidth;
      }

      strokePath(ctx, points);
      ctx.stroke();

      if (isMarker) {
        ctx.globalAlpha = 1;
        ctx.globalCompositeOperation = 'source-over';
      }

      if (isSelected) {
        const bounds = getStrokeBounds(stroke.points, pageNum);
        if (bounds) {
          ctx.setLineDash([10, 8]);
          ctx.strokeStyle = 'rgba(25, 118, 210, 0.85)';
          ctx.lineWidth = 2;
          ctx.strokeRect(
            bounds.x - baseLineWidth,
            bounds.y - baseLineWidth,
            bounds.width + baseLineWidth * 2,
            bounds.height + baseLineWidth * 2
          );
        }
      }

      ctx.restore();
    }

    function strokePath(ctx, points) {
      ctx.beginPath();

      if (points.length === 1) {
        ctx.arc(points[0].x, points[0].y, Math.max(ctx.lineWidth / 2, 1), 0, Math.PI * 2);
        return;
      }

      ctx.moveTo(points[0].x, points[0].y);

      if (points.length === 2) {
        ctx.lineTo(points[1].x, points[1].y);
        return;
      }

      for (let index = 1; index < points.length - 1; index += 1) {
        const current = points[index];
        const next = points[index + 1];
        const midX = (current.x + next.x) / 2;
        const midY = (current.y + next.y) / 2;
        ctx.quadraticCurveTo(current.x, current.y, midX, midY);
      }

      const last = points[points.length - 1];
      ctx.lineTo(last.x, last.y);
    }

    function redrawDrawings(pageNum, previewStroke) {
      const canvas = document.querySelector('.draw-canvas[data-page="' + pageNum + '"]');
      if (!canvas) {
        return;
      }

      const ctx = canvas.getContext('2d');
      ctx.clearRect(0, 0, canvas.width, canvas.height);

      const pageAnnotations = ensurePageAnnotations(pageNum);
      pageAnnotations.drawings.forEach((stroke) => {
        const isSelected =
          !!selectedAnnotation &&
          selectedAnnotation.type === 'drawing' &&
          selectedAnnotation.page === Number(pageNum) &&
          selectedAnnotation.id === stroke.id;
        drawStrokeToCanvas(ctx, stroke, pageNum, { selected: isSelected });
      });

      if (previewStroke) {
        drawStrokeToCanvas(ctx, previewStroke, pageNum, { selected: false });
      }
    }

    function getNormalizedEventPoint(target, event) {
      const coords = getEventCoords(event);
      const bounds = target.getBoundingClientRect();
      return {
        x: clamp((coords.clientX - bounds.left) / bounds.width, 0, 1),
        y: clamp((coords.clientY - bounds.top) / bounds.height, 0, 1),
      };
    }

    function setupAnnotationLayer(overlay, pageNum) {
      let draftRect = null;
      let startPoint = null;

      function onDown(event) {
        if (
          currentMode === 'view' &&
          selectedAnnotation &&
          selectedAnnotation.type === 'drawing' &&
          selectedAnnotation.page === Number(pageNum)
        ) {
          const point = getNormalizedEventPoint(overlay, event);
          const drawing = findDrawingAtPoint(pageNum, point);

          if (drawing && drawing.id === selectedAnnotation.id) {
            event.preventDefault();
            event.stopPropagation();
            beginDrag(
              {
                kind: 'drawing',
                annotation: drawing,
                pageNum,
                originalPoints: drawing.points.map((dragPoint) => ({
                  x: dragPoint.x,
                  y: dragPoint.y,
                })),
              },
              event
            );
            return;
          }
        }

        if (currentMode === 'view') {
          if (event.target === overlay) {
            const now = Date.now();
            if (
              crossPdfLastTap.page === pageNum &&
              now - crossPdfLastTap.time < 450
            ) {
              crossPdfLastTap = { time: 0, page: 0 };
              event.preventDefault();
              event.stopPropagation();
              suppressNextClick = true;
              const c = getEventCoords(event);
              beginPageGhostDrag(pageNum, c.clientX, c.clientY);
              return;
            }
            crossPdfLastTap = { time: now, page: pageNum };
          }
          return;
        }

        if (currentMode !== 'highlight' && currentMode !== 'text') {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        startPoint = getNormalizedEventPoint(overlay, event);

        if (currentMode === 'highlight') {
          draftRect = document.createElement('div');
          draftRect.className = 'highlight-rect';
          draftRect.style.background = highlightColor;
          overlay.appendChild(draftRect);
          updateDraftRect(draftRect, startPoint, startPoint, pageNum);
        } else if (currentMode === 'text') {
          sendMessage({
            type: 'requestTextInput',
            page: pageNum,
            x: startPoint.x,
            y: startPoint.y,
          });
        }
      }

      function onMove(event) {
        if (currentMode === 'view') {
          return;
        }

        if (currentMode !== 'highlight' || !draftRect || !startPoint) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();
        const currentPoint = getNormalizedEventPoint(overlay, event);
        updateDraftRect(draftRect, startPoint, currentPoint, pageNum);
      }

      function onUp(event) {
        if (currentMode === 'view') {
          return;
        }

        if (currentMode !== 'highlight' || !draftRect || !startPoint) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        const currentPoint = getNormalizedEventPoint(overlay, event);
        const annotation = finalizeHighlight(startPoint, currentPoint);

        if (!annotation) {
          draftRect.remove();
          draftRect = null;
          startPoint = null;
          return;
        }

        annotation.id = createAnnotationId('highlight');
        annotation.color = highlightColor;

        const pageAnnotations = ensurePageAnnotations(pageNum);
        pageAnnotations.highlights.push(annotation);
        draftRect.replaceWith(createHighlightElement(annotation, pageNum));

        draftRect = null;
        startPoint = null;
        persistAnnotations();
      }

      overlay.addEventListener('touchstart', onDown, { passive: false });
      overlay.addEventListener('touchmove', onMove, { passive: false });
      overlay.addEventListener('touchend', onUp, { passive: false });
      overlay.addEventListener('touchcancel', function () {});
      overlay.addEventListener('mousedown', onDown);
      overlay.addEventListener('mousemove', onMove);
      overlay.addEventListener('mouseup', onUp);
      overlay.addEventListener('mouseleave', function () {});

      overlay.addEventListener('click', (event) => {
        if (currentMode !== 'view' || event.target !== overlay) {
          return;
        }
        if (suppressNextClick) {
          suppressNextClick = false;
          event.preventDefault();
          event.stopPropagation();
          return;
        }

        const point = getNormalizedEventPoint(overlay, event);
        const drawing = findDrawingAtPoint(pageNum, point);

        if (drawing) {
          selectAnnotation({
            id: drawing.id,
            type: 'drawing',
            page: pageNum,
          });
        } else {
          selectAnnotation(null);
          setSelectedPage(pageNum);
        }
      });
    }

    function updateDraftRect(rect, startPoint, endPoint, pageNum) {
      const metric = getPageMetric(pageNum);
      if (!metric) {
        return;
      }

      const x = Math.min(startPoint.x, endPoint.x);
      const y = Math.min(startPoint.y, endPoint.y);
      const width = Math.abs(endPoint.x - startPoint.x);
      const height = Math.abs(endPoint.y - startPoint.y);

      rect.style.left = x * metric.displayWidth + 'px';
      rect.style.top = y * metric.displayHeight + 'px';
      rect.style.width = width * metric.displayWidth + 'px';
      rect.style.height = height * metric.displayHeight + 'px';
    }

    function finalizeHighlight(startPoint, endPoint) {
      const x = Math.min(startPoint.x, endPoint.x);
      const y = Math.min(startPoint.y, endPoint.y);
      const width = Math.abs(endPoint.x - startPoint.x);
      const height = Math.abs(endPoint.y - startPoint.y);

      if (width < 0.012 && height < 0.012) {
        return null;
      }

      return { x, y, width, height };
    }

    function addTextNote(pageNum, x, y, text) {
      const pageAnnotations = ensurePageAnnotations(pageNum);
      pageAnnotations.notes.push({
        id: createAnnotationId('note'),
        x: clamp(safeNumber(x, 0), 0, 1),
        y: clamp(safeNumber(y, 0), 0, 1),
        text: String(text || ''),
      });
      renderPageAnnotations(pageNum);
      persistAnnotations();
    }

    function setupDrawCanvas(canvas, pageNum) {
      let drawing = false;
      let previewStroke = null;

      function onDrawStart(event) {
        if (currentMode !== 'draw' && currentMode !== 'marker') {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        drawing = true;
        const bounds = canvas.getBoundingClientRect();
        if (currentMode === 'marker') {
          previewStroke = {
            id: createAnnotationId('drawing'),
            kind: 'marker',
            color: highlightColor,
            width: clamp(markerBrushPx / bounds.width, 0.002, 0.09),
            points: [getNormalizedEventPoint(canvas, event)],
          };
        } else {
          previewStroke = {
            id: createAnnotationId('drawing'),
            kind: 'pen',
            color: drawColor,
            width: clamp(drawWidth / bounds.width, 0.001, 0.05),
            points: [getNormalizedEventPoint(canvas, event)],
          };
        }

        redrawDrawings(pageNum, previewStroke);
      }

      function onDrawMove(event) {
        if (
          !drawing ||
          (currentMode !== 'draw' && currentMode !== 'marker') ||
          !previewStroke
        ) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        const nextPoint = getNormalizedEventPoint(canvas, event);
        const lastPoint = previewStroke.points[previewStroke.points.length - 1];
        const deltaX = nextPoint.x - lastPoint.x;
        const deltaY = nextPoint.y - lastPoint.y;

        if (Math.sqrt(deltaX * deltaX + deltaY * deltaY) < 0.0015) {
          return;
        }

        previewStroke.points.push(nextPoint);
        redrawDrawings(pageNum, previewStroke);
      }

      function onDrawEnd() {
        if (!drawing || !previewStroke) {
          return;
        }

        drawing = false;
        const pageAnnotations = ensurePageAnnotations(pageNum);

        if (previewStroke.points.length === 1) {
          previewStroke.points.push({
            x: clamp(previewStroke.points[0].x + 0.0001, 0, 1),
            y: previewStroke.points[0].y,
          });
        }

        pageAnnotations.drawings.push(previewStroke);
        redrawDrawings(pageNum);
        previewStroke = null;
        persistAnnotations();
      }

      canvas.addEventListener('touchstart', onDrawStart, { passive: false });
      canvas.addEventListener('touchmove', onDrawMove, { passive: false });
      canvas.addEventListener('touchend', onDrawEnd, { passive: false });
      canvas.addEventListener('touchcancel', onDrawEnd, { passive: false });
      canvas.addEventListener('mousedown', onDrawStart);
      canvas.addEventListener('mousemove', onDrawMove);
      canvas.addEventListener('mouseup', onDrawEnd);
      canvas.addEventListener('mouseleave', onDrawEnd);
    }

    function getEventCoords(event) {
      if (event.touches && event.touches.length > 0) {
        return {
          clientX: event.touches[0].clientX,
          clientY: event.touches[0].clientY,
        };
      }

      if (event.changedTouches && event.changedTouches.length > 0) {
        return {
          clientX: event.changedTouches[0].clientX,
          clientY: event.changedTouches[0].clientY,
        };
      }

      return {
        clientX: event.clientX,
        clientY: event.clientY,
      };
    }

    function setupCrossPaneDragHandle(handle, pageNum) {
      function onTap(event) {
        if (!CROSS_PANE_DRAG) {
          return;
        }
        event.stopPropagation();
        event.preventDefault();
        sendMessage({
          type: 'crossPdfPickSource',
          slot: VIEWER_SLOT,
          page: pageNum,
        });
      }
      handle.addEventListener('click', onTap);
    }

    function isNearCrossPaneEdge(clientX, clientY, vw, vh) {
      const margin = Math.min(vw, vh) * 0.17;
      if (DUAL_LAYOUT === 'column') {
        if (VIEWER_SLOT === 1) {
          return clientY > vh - margin;
        }
        return clientY < margin;
      }
      if (VIEWER_SLOT === 1) {
        return clientX > vw - margin;
      }
      return clientX < margin;
    }

    function closePageMenu() {
      const back = document.getElementById('page-menu-backdrop');
      const drop = document.getElementById('page-menu-dropdown');
      if (back) {
        back.remove();
      }
      if (drop) {
        drop.remove();
      }
      pageMenuOpenFor = 0;
    }

    function openPageMenu(pageNum, anchorEl) {
      closePageMenu();
      pageMenuOpenFor = pageNum;
      const rect = anchorEl.getBoundingClientRect();
      const backdrop = document.createElement('div');
      backdrop.id = 'page-menu-backdrop';
      backdrop.className = 'page-menu-backdrop';
      backdrop.addEventListener('click', closePageMenu);
      const menu = document.createElement('div');
      menu.id = 'page-menu-dropdown';
      menu.className = 'page-menu-dropdown';
      const editBtn = document.createElement('button');
      editBtn.type = 'button';
      editBtn.className = 'page-menu-item';
      editBtn.textContent = 'Edit pages (order, remove…)';
      editBtn.addEventListener('click', function () {
        closePageMenu();
        sendMessage({
          type: 'pageMenuAction',
          action: 'editPages',
          slot: VIEWER_SLOT,
          page: pageNum,
        });
      });
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'page-menu-item danger';
      removeBtn.textContent = 'Remove this page…';
      removeBtn.addEventListener('click', function () {
        closePageMenu();
        sendMessage({
          type: 'pageMenuAction',
          action: 'removePage',
          slot: VIEWER_SLOT,
          page: pageNum,
        });
      });
      menu.appendChild(editBtn);
      menu.appendChild(removeBtn);
      document.body.appendChild(backdrop);
      document.body.appendChild(menu);
      const mw = menu.offsetWidth || 200;
      const mh = menu.offsetHeight || 120;
      let left = rect.left + rect.width / 2 - mw / 2;
      let top = rect.bottom + 6;
      if (left < 8) {
        left = 8;
      }
      if (left + mw > window.innerWidth - 8) {
        left = window.innerWidth - mw - 8;
      }
      if (top + mh > window.innerHeight - 8) {
        top = rect.top - mh - 6;
      }
      menu.style.left = left + 'px';
      menu.style.top = top + 'px';
    }

    function setupPageMenuButton(btn, pageNum) {
      btn.addEventListener('click', function (event) {
        event.stopPropagation();
        event.preventDefault();
        if (pageMenuOpenFor === pageNum) {
          closePageMenu();
        } else {
          openPageMenu(pageNum, btn);
        }
      });
    }

    function finishPageGhostDrag(fromPage, clientX, clientY) {
      const vw = window.innerWidth;
      const vh = window.innerHeight;
      if (CROSS_PANE_DRAG && isNearCrossPaneEdge(clientX, clientY, vw, vh)) {
        sendMessage({
          type: 'crossPdfPickSource',
          slot: VIEWER_SLOT,
          page: fromPage,
        });
        return;
      }
      let el = document.elementFromPoint(clientX, clientY);
      let targetPage = null;
      let depth = 0;
      while (el && depth < 22) {
        if (el.classList && el.classList.contains('page-wrapper')) {
          targetPage = Number(el.dataset.page);
          break;
        }
        el = el.parentElement;
        depth += 1;
      }
      if (
        targetPage &&
        targetPage !== fromPage &&
        totalPages > 1
      ) {
        sendMessage({
          type: 'pageReorderDrag',
          slot: VIEWER_SLOT,
          fromPage: fromPage,
          toAfterPage: targetPage,
        });
      }
    }

    function beginPageGhostDrag(pageNum, clientX, clientY) {
      const wrap = document.querySelector('.page-wrapper[data-page="' + pageNum + '"]');
      if (!wrap) {
        return;
      }
      const rect = wrap.getBoundingClientRect();
      const ghost = wrap.cloneNode(true);
      ghost.classList.add('page-drag-ghost');
      ghost.style.position = 'fixed';
      ghost.style.zIndex = '100000';
      ghost.style.pointerEvents = 'none';
      ghost.style.left = clientX - rect.width / 2 + 'px';
      ghost.style.top = clientY - rect.height / 2 + 'px';
      ghost.style.width = rect.width + 'px';
      ghost.style.height = rect.height + 'px';
      ghost.style.transform = 'scale(1.06)';
      ghost.style.boxShadow = '0 22px 48px rgba(0, 0, 0, 0.48)';
      document.body.appendChild(ghost);
      wrap.classList.add('page-wrapper-dimmed');
      const gw = rect.width;
      const gh = rect.height;
      pageGhostDrag = { pageNum: pageNum, ghost: ghost, wrap: wrap, gw: gw, gh: gh };

      function move(ev) {
        if (!pageGhostDrag) {
          return;
        }
        ev.preventDefault();
        const c = getEventCoords(ev);
        pageGhostDrag.ghost.style.left = c.clientX - pageGhostDrag.gw / 2 + 'px';
        pageGhostDrag.ghost.style.top = c.clientY - pageGhostDrag.gh / 2 + 'px';
      }

      function end(ev) {
        document.removeEventListener('touchmove', move);
        document.removeEventListener('mousemove', move);
        document.removeEventListener('touchend', end);
        document.removeEventListener('mouseup', end);
        const c = getEventCoords(ev);
        const px = c.clientX;
        const py = c.clientY;
        if (pageGhostDrag && pageGhostDrag.ghost) {
          pageGhostDrag.ghost.remove();
        }
        if (pageGhostDrag && pageGhostDrag.wrap) {
          pageGhostDrag.wrap.classList.remove('page-wrapper-dimmed');
        }
        const fromP = pageGhostDrag ? pageGhostDrag.pageNum : pageNum;
        pageGhostDrag = null;
        finishPageGhostDrag(fromP, px, py);
      }

      document.addEventListener('touchmove', move, { passive: false });
      document.addEventListener('mousemove', move);
      document.addEventListener('touchend', end);
      document.addEventListener('mouseup', end);
    }

    document.addEventListener('touchmove', updateActiveDrag, { passive: false });
    document.addEventListener('touchend', endActiveDrag, { passive: false });
    document.addEventListener('touchcancel', endActiveDrag, { passive: false });
    document.addEventListener('mousemove', updateActiveDrag);
    document.addEventListener('mouseup', endActiveDrag);

    function updateInteractionMode() {
      document.querySelectorAll('.annotation-layer').forEach((element) => {
        element.classList.toggle(
          'active',
          currentMode !== 'draw' && currentMode !== 'marker'
        );
        element.classList.toggle('editing', currentMode === 'highlight' || currentMode === 'text');
      });

      document.querySelectorAll('.draw-canvas').forEach((element) => {
        element.classList.toggle('active', currentMode === 'draw' || currentMode === 'marker');
      });
    }

    function setMode(mode) {
      currentMode = mode;
      if (mode !== 'view') {
        selectAnnotation(null);
      }
      updateInteractionMode();
      sendMessage({ type: 'modeChanged', mode });
    }

    function setDrawColor(color) {
      drawColor = color;
    }

    function setHighlightColor(color) {
      highlightColor = color;
    }

    function setDrawBrushWidth(px) {
      const next = Number(px);
      if (Number.isFinite(next)) {
        drawWidth = clamp(next, 1.5, 28);
      }
    }

    function setMarkerBrushWidth(px) {
      const next = Number(px);
      if (Number.isFinite(next)) {
        markerBrushPx = clamp(next, 6, 40);
      }
    }

    function deleteSelectedAnnotation() {
      if (!selectedAnnotation) {
        return;
      }

      const pageNum = Number(selectedAnnotation.page);
      const pageAnnotations = ensurePageAnnotations(pageNum);

      if (selectedAnnotation.type === 'highlight') {
        pageAnnotations.highlights = pageAnnotations.highlights.filter(
          (annotation) => annotation.id !== selectedAnnotation.id
        );
      } else if (selectedAnnotation.type === 'note') {
        pageAnnotations.notes = pageAnnotations.notes.filter(
          (annotation) => annotation.id !== selectedAnnotation.id
        );
      } else if (selectedAnnotation.type === 'drawing') {
        pageAnnotations.drawings = pageAnnotations.drawings.filter(
          (annotation) => annotation.id !== selectedAnnotation.id
        );
      }

      renderPageAnnotations(pageNum);
      selectAnnotation(null);
      persistAnnotations();
    }

    function clearAnnotations(pageNum, shouldPersist) {
      annotations[pageNum] = createEmptyPageAnnotations();
      renderPageAnnotations(pageNum);

      if (selectedAnnotation && Number(selectedAnnotation.page) === Number(pageNum)) {
        selectAnnotation(null);
      }

      if (shouldPersist !== false) {
        persistAnnotations();
      }
    }

    function clearAllAnnotations() {
      for (let pageNum = 1; pageNum <= totalPages; pageNum += 1) {
        annotations[pageNum] = createEmptyPageAnnotations();
        renderPageAnnotations(pageNum);
      }
      selectAnnotation(null);
      persistAnnotations();
    }

    async function loadPdf() {
      try {
        const base64 = "${base64Data}";
        const raw = atob(base64);
        const uint8 = new Uint8Array(raw.length);
        for (let index = 0; index < raw.length; index += 1) {
          uint8[index] = raw.charCodeAt(index);
        }

        pdfDoc = await pdfjsLib.getDocument({ data: uint8 }).promise;
        totalPages = pdfDoc.numPages;

        for (let pageNum = 1; pageNum <= totalPages; pageNum += 1) {
          await renderPage(pageNum);
        }

        document.getElementById('loading').style.display = 'none';
        updateInteractionMode();
        setSelectedPage(selectedPage, true);
        sendMessage({ type: 'pdfLoaded', totalPages });
      } catch (error) {
        document.getElementById('loading').style.display = 'none';
        const element = document.getElementById('error-msg');
        element.textContent = 'Failed to load PDF: ' + error.message;
        element.style.display = 'block';
        sendMessage({ type: 'error', message: error.message });
      }
    }

    async function renderPage(pageNum) {
      const page = await pdfDoc.getPage(pageNum);
      const fitViewport = page.getViewport({ scale: 1 });
      const availableWidth = Math.max(window.innerWidth - 24, 260);
      const displayScale = availableWidth / fitViewport.width;
      const deviceScale = Math.max(window.devicePixelRatio || 1, 1.5);
      const renderScale = displayScale * deviceScale;
      const renderViewport = page.getViewport({ scale: renderScale });
      const displayWidth = fitViewport.width * displayScale;
      const displayHeight = fitViewport.height * displayScale;

      const wrapper = document.createElement('div');
      wrapper.className = 'page-wrapper';
      wrapper.dataset.page = pageNum;
      wrapper.style.width = displayWidth + 'px';
      wrapper.style.height = displayHeight + 'px';

      const canvas = document.createElement('canvas');
      canvas.width = renderViewport.width;
      canvas.height = renderViewport.height;
      canvas.style.width = displayWidth + 'px';
      canvas.style.height = displayHeight + 'px';

      const canvasContext = canvas.getContext('2d');
      await page.render({ canvasContext, viewport: renderViewport }).promise;
      wrapper.appendChild(canvas);

      const drawCanvas = document.createElement('canvas');
      drawCanvas.className = 'draw-canvas';
      drawCanvas.width = Math.round(displayWidth * deviceScale);
      drawCanvas.height = Math.round(displayHeight * deviceScale);
      drawCanvas.style.width = displayWidth + 'px';
      drawCanvas.style.height = displayHeight + 'px';
      drawCanvas.dataset.page = pageNum;
      wrapper.appendChild(drawCanvas);

      const overlay = document.createElement('div');
      overlay.className = 'annotation-layer';
      overlay.dataset.page = pageNum;
      wrapper.appendChild(overlay);

      const label = document.createElement('div');
      label.dataset.page = pageNum;
      label.className = 'page-label page-label-row';
      const text = document.createElement('span');
      text.className = 'page-label-text';
      text.textContent = 'Page ' + pageNum + ' of ' + totalPages;
      label.appendChild(text);
      if (CROSS_PANE_DRAG) {
        const dragHandle = document.createElement('button');
        dragHandle.type = 'button';
        dragHandle.className = 'page-drag-handle';
        dragHandle.setAttribute(
          'aria-label',
          'Tap to copy or move this page into the other PDF'
        );
        dragHandle.textContent = '⇄';
        setupCrossPaneDragHandle(dragHandle, pageNum);
        label.appendChild(dragHandle);
      }
      const menuBtn = document.createElement('button');
      menuBtn.type = 'button';
      menuBtn.className = 'page-menu-btn';
      menuBtn.setAttribute('aria-label', 'Page options');
      menuBtn.textContent = '⋮';
      setupPageMenuButton(menuBtn, pageNum);
      label.appendChild(menuBtn);

      const container = document.getElementById('pdf-container');
      container.appendChild(wrapper);
      container.appendChild(label);

      pageMetrics[pageNum] = {
        displayWidth,
        displayHeight,
        drawCanvasWidth: drawCanvas.width,
        drawCanvasHeight: drawCanvas.height,
      };

      ensurePageAnnotations(pageNum);
      setupDrawCanvas(drawCanvas, pageNum);
      setupAnnotationLayer(overlay, pageNum);
      renderPageAnnotations(pageNum);
    }

    function sendMessage(data) {
      if (window.ReactNativeWebView) {
        window.ReactNativeWebView.postMessage(JSON.stringify(data));
      }
    }

    function wrapTextLines(ctx, text, maxWidth) {
      const words = String(text || '')
        .split(/\s+/)
        .filter(Boolean);
      const lines = [];
      let line = '';
      words.forEach((word) => {
        const test = line ? line + ' ' + word : word;
        if (ctx.measureText(test).width > maxWidth && line) {
          lines.push(line);
          line = word;
        } else {
          line = test;
        }
      });
      if (line) {
        lines.push(line);
      }
      return lines.length ? lines : [''];
    }

    function exportPageSnapshot(requestId, pageNum, mimeType, quality) {
      const rid = String(requestId || '');
      const page = clamp(safeNumber(pageNum, 0), 1, totalPages || 1);
      const wrapper = document.querySelector('.page-wrapper[data-page="' + page + '"]');
      if (!wrapper || !totalPages) {
        sendMessage({
          type: 'pageExportImageError',
          requestId: rid,
          message: 'Page is not available yet.',
        });
        return;
      }

      const canvases = wrapper.querySelectorAll('canvas');
      const pdfCanvas = canvases[0];
      const drawCanvas = canvases[1];
      if (!pdfCanvas) {
        sendMessage({
          type: 'pageExportImageError',
          requestId: rid,
          message: 'Page canvas is not ready.',
        });
        return;
      }

      const w = pdfCanvas.width;
      const h = pdfCanvas.height;
      const exportCanvas = document.createElement('canvas');
      exportCanvas.width = w;
      exportCanvas.height = h;
      const ctx = exportCanvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(pdfCanvas, 0, 0);

      if (drawCanvas) {
        ctx.drawImage(drawCanvas, 0, 0, drawCanvas.width, drawCanvas.height, 0, 0, w, h);
      }

      const pageAnnotations = ensurePageAnnotations(page);
      pageAnnotations.highlights.forEach((ann) => {
        ctx.fillStyle = ann.color || highlightColor;
        ctx.globalAlpha = 0.38;
        ctx.fillRect(ann.x * w, ann.y * h, ann.width * w, ann.height * h);
        ctx.globalAlpha = 1;
      });

      const fontSize = Math.max(12, Math.round(w * 0.022));
      ctx.textBaseline = 'top';
      pageAnnotations.notes.forEach((ann) => {
        const maxTextW = Math.min(w * 0.45, 220 * (w / 420));
        ctx.font = fontSize + 'px -apple-system, BlinkMacSystemFont, sans-serif';
        const lines = wrapTextLines(ctx, ann.text, maxTextW);
        const lineHeight = Math.round(fontSize * 1.38);
        const pad = Math.max(6, Math.round(w * 0.014));
        const boxW = Math.min(maxTextW + pad * 2, w - ann.x * w - 4);
        const boxH = lines.length * lineHeight + pad * 2;
        const nx = ann.x * w;
        const ny = ann.y * h;
        ctx.fillStyle = 'rgba(255, 249, 196, 0.96)';
        ctx.strokeStyle = 'rgba(249, 168, 37, 0.85)';
        ctx.lineWidth = Math.max(1, w / 520);
        const r = Math.min(10, pad);
        ctx.beginPath();
        ctx.moveTo(nx + r, ny);
        ctx.arcTo(nx + boxW, ny, nx + boxW, ny + boxH, r);
        ctx.arcTo(nx + boxW, ny + boxH, nx, ny + boxH, r);
        ctx.arcTo(nx, ny + boxH, nx, ny, r);
        ctx.arcTo(nx, ny, nx + boxW, ny, r);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.fillStyle = '#333333';
        lines.forEach((line, index) => {
          ctx.fillText(line, nx + pad, ny + pad + index * lineHeight);
        });
      });

      const mime = typeof mimeType === 'string' && mimeType ? mimeType : 'image/png';
      const q =
        typeof quality === 'number' && !Number.isNaN(quality)
          ? clamp(quality, 0.5, 1)
          : 0.92;
      let dataUrl = '';
      try {
        if (mime === 'image/jpeg' || mime === 'image/jpg') {
          dataUrl = exportCanvas.toDataURL('image/jpeg', q);
        } else if (mime === 'image/webp') {
          dataUrl = exportCanvas.toDataURL('image/webp', q);
        } else {
          dataUrl = exportCanvas.toDataURL('image/png');
        }
      } catch (err) {
        sendMessage({
          type: 'pageExportImageError',
          requestId: rid,
          message: err && err.message ? err.message : 'Could not encode image.',
        });
        return;
      }

      const base64Payload = (dataUrl.split(',')[1] || '').trim();
      if (!base64Payload) {
        sendMessage({
          type: 'pageExportImageError',
          requestId: rid,
          message: 'Image export returned empty data.',
        });
        return;
      }

      sendMessage({
        type: 'pageExportImageReady',
        requestId: rid,
        page,
        mimeType: mime === 'image/jpg' ? 'image/jpeg' : mime,
        base64: base64Payload,
      });
    }

    function dispatchViewerCommand(parsed) {
      if (!parsed || typeof parsed.command !== 'string') {
        return;
      }
      switch (parsed.command) {
        case 'setMode':
          setMode(parsed.mode);
          break;
        case 'setDrawColor':
          setDrawColor(parsed.color);
          break;
        case 'setHighlightColor':
          setHighlightColor(parsed.color);
          break;
        case 'setDrawBrushWidth':
          setDrawBrushWidth(parsed.px);
          break;
        case 'setMarkerBrushWidth':
          setMarkerBrushWidth(parsed.px);
          break;
        case 'addTextNote':
          addTextNote(parsed.page, parsed.x, parsed.y, parsed.text);
          break;
        case 'clearPage':
          clearAnnotations(parsed.page);
          break;
        case 'clearAll':
          clearAllAnnotations();
          break;
        case 'deleteSelected':
          deleteSelectedAnnotation();
          break;
        case 'exportPageImage':
          exportPageSnapshot(parsed.requestId, parsed.page, parsed.mimeType, parsed.quality);
          break;
        default:
          break;
      }
    }

    window.dispatchViewerCommand = dispatchViewerCommand;

    function handleIncomingMessage(message) {
      try {
        const parsed = JSON.parse(message);
        dispatchViewerCommand(parsed);
      } catch (_) {}
    }

    window.addEventListener('message', (event) => handleIncomingMessage(event.data));
    document.addEventListener('message', (event) => handleIncomingMessage(event.data));

    loadPdf();
  </script>
</body>
</html>
  `.trim();
}
