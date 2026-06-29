const $ = (id) => document.getElementById(id);

const fields = {
  left: $("left"),
  top: $("top"),
  width: $("width"),
  height: $("height"),
  opacity: $("opacity"),
  fillColor: $("fillColor"),
  fillHex: $("fillHex"),
  fillOpacity: $("fillOpacity"),
  lineColor: $("lineColor"),
  lineHex: $("lineHex"),
  lineOpacity: $("lineOpacity"),
  horizontalSpacing: $("horizontalSpacing"),
  verticalSpacing: $("verticalSpacing"),
  lineWeight: $("lineWeight")
};

let selectedShapeId = null;
let selectedShapeIds = [];
let ratioLocked = false;
let aspectRatio = 1;
let refreshTimer = null;
let applyTimer = null;
let liveRefreshTimer = null;
let isApplying = false;
let isRefreshing = false;
let pendingApplyFields = new Set();
let latestSpacing = { horizontal: null, vertical: null };
let currentBounds = { left: 0, top: 0, width: 1, height: 1 };

const demoShape = {
  id: "demo",
  name: "Rectangle 1",
  left: 120,
  top: 86,
  width: 360,
  height: 204,
  visible: true,
  fillColor: "FFFFFF",
  fillOpacity: 100,
  lineColor: "0D0D0D",
  lineOpacity: 100,
  lineWeight: 1
};

function isOfficeReady() {
  return Boolean(window.Office && window.PowerPoint);
}

function setStatus(message) {
  $("status").textContent = message;
}

function isEditingFormField() {
  return ["INPUT", "SELECT", "TEXTAREA"].includes(document.activeElement?.tagName);
}

function formatError(error) {
  const debug = error?.debugInfo ? ` ${JSON.stringify(error.debugInfo)}` : "";
  return `${error?.message || String(error)}${debug}`;
}

function clamp(value, min, max) {
  const number = Number(value);
  if (Number.isNaN(number)) return min;
  return Math.min(max, Math.max(min, number));
}

function cleanHex(value, fallback = "FFFFFF") {
  const normalized = String(value || "").replace(/[^0-9a-f]/gi, "").slice(0, 6).toUpperCase();
  return normalized.length === 6 ? normalized : fallback;
}

function hexToColorInput(hex) {
  return `#${cleanHex(hex)}`;
}

function hexToOfficeColor(hex) {
  return `#${cleanHex(hex)}`;
}

function syncColorPair(colorInput, hexInput, nextHex) {
  const hex = cleanHex(nextHex, hexInput.value || "FFFFFF");
  colorInput.value = hexToColorInput(hex);
  hexInput.value = hex;
}

function boundsForShapes(shapes) {
  const left = Math.min(...shapes.map((shape) => shape.left));
  const top = Math.min(...shapes.map((shape) => shape.top));
  const right = Math.max(...shapes.map((shape) => shape.left + shape.width));
  const bottom = Math.max(...shapes.map((shape) => shape.top + shape.height));
  return {
    left,
    top,
    width: right - left,
    height: bottom - top,
    right,
    bottom
  };
}

function commonOrFirst(shapes, getter, fallback) {
  if (!shapes.length) return fallback;
  const first = getter(shapes[0]);
  return shapes.every((shape) => getter(shape) === first) ? first : first ?? fallback;
}

function commonRoundedOrMixed(shapes, property) {
  if (!shapes.length) return "";
  const values = shapes.map((shape) => Math.round(shape[property] || 0));
  const first = values[0];
  return values.every((value) => value === first) ? String(first) : "混合";
}

function parseNumberField(field, fallback, minimum = -Infinity) {
  const parsed = Number(String(field.value).replace(/[^\d.-]/g, ""));
  const number = Number.isFinite(parsed) ? parsed : fallback;
  return Math.max(minimum, number);
}

function isFiniteNumberText(value) {
  return /^-?(?:\d+|\d*\.\d+)$/.test(String(value).trim());
}

function restoreInputValue(field) {
  field.value = field.dataset.entryValue ?? field.defaultValue ?? "";
}

function numericValueForStep(field) {
  return isFiniteNumberText(field.value) ? Number(field.value) : 0;
}

function stepInputValue(field, direction, options = {}, event = {}) {
  const step = options.step || 1;
  const multiplier = event.shiftKey ? 10 : 1;
  const stepped = numericValueForStep(field) + direction * step * multiplier;
  const next = clamp(stepped, options.min ?? -Infinity, options.max ?? Infinity);
  field.value = Number.isInteger(next) ? String(next) : String(Number(next.toFixed(2)));
  field.dataset.entryValue = field.value;
}

function selectInputContents(field) {
  window.setTimeout(() => {
    if (document.activeElement === field) field.select();
  }, 0);
}

function bindSelectableInput(field, options = {}) {
  const validate = options.validate || isFiniteNumberText;
  let focusSelected = false;

  field.addEventListener("focus", () => {
    field.dataset.entryValue = field.value;
    focusSelected = true;
    selectInputContents(field);
  });

  field.addEventListener("click", () => {
    selectInputContents(field);
  });

  field.addEventListener("mouseup", (event) => {
    if (!focusSelected) return;
    event.preventDefault();
    focusSelected = false;
  });

  field.addEventListener("input", () => {
    if (!validate(field.value)) {
      pendingApplyFields.clear();
      window.clearTimeout(applyTimer);
      return;
    }
    options.onValid?.();
  });

  field.addEventListener("keydown", (event) => {
    if (options.allowStep !== false && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
      event.preventDefault();
      stepInputValue(field, event.key === "ArrowUp" ? 1 : -1, options, event);
      options.onValid?.();
      selectInputContents(field);
      return;
    }

    if (event.key !== "Enter") return;
    event.preventDefault();
    if (validate(field.value)) {
      field.dataset.entryValue = field.value;
      options.onValid?.();
      field.blur();
      return;
    }
    restoreInputValue(field);
    field.blur();
  });

  field.addEventListener("blur", () => {
    if (validate(field.value)) {
      field.dataset.entryValue = field.value;
      return;
    }
    restoreInputValue(field);
  });
}

function bindTextInputSelection(field) {
  field.addEventListener("focus", () => {
    field.dataset.entryValue = field.value;
    selectInputContents(field);
  });
  field.addEventListener("click", () => {
    selectInputContents(field);
  });
}

function centerX(shape) {
  return shape.left + shape.width / 2;
}

function centerY(shape) {
  return shape.top + shape.height / 2;
}

function overlapLength(aStart, aEnd, bStart, bEnd) {
  return Math.max(0, Math.min(aEnd, bEnd) - Math.max(aStart, bStart));
}

function haveVerticalOverlap(a, b) {
  const overlap = overlapLength(a.top, a.top + a.height, b.top, b.top + b.height);
  return overlap >= Math.min(a.height, b.height) * 0.35;
}

function haveHorizontalOverlap(a, b) {
  const overlap = overlapLength(a.left, a.left + a.width, b.left, b.left + b.width);
  return overlap >= Math.min(a.width, b.width) * 0.35;
}

function groupByAxis(shapes, axis) {
  const sorted = [...shapes].sort((a, b) => (axis === "row" ? centerY(a) - centerY(b) : centerX(a) - centerX(b)));
  const groups = [];

  sorted.forEach((shape) => {
    const group = groups.find((items) => {
      const last = items[items.length - 1];
      return axis === "row" ? haveVerticalOverlap(last, shape) : haveHorizontalOverlap(last, shape);
    });
    if (group) group.push(shape);
    else groups.push([shape]);
  });

  return groups;
}

function gapsForGroups(groups, axis) {
  return groups.flatMap((group) => {
    if (group.length < 2) return [];
    const sorted = [...group].sort((a, b) => (axis === "horizontal" ? a.left - b.left : a.top - b.top));
    return sorted.slice(1).map((shape, index) => {
      const previous = sorted[index];
      return axis === "horizontal"
        ? shape.left - (previous.left + previous.width)
        : shape.top - (previous.top + previous.height);
    });
  });
}

function summarizeGaps(gaps) {
  if (!gaps.length) return null;
  const rounded = gaps.map((gap) => Math.round(gap));
  const first = rounded[0];
  return rounded.every((gap) => gap === first) ? String(first) : "混在";
}

function spacingInfoForShapes(shapes) {
  if (shapes.length < 2) return { horizontal: null, vertical: null };
  const rowGroups = groupByAxis(shapes, "row").filter((group) => group.length > 1);
  const columnGroups = groupByAxis(shapes, "column").filter((group) => group.length > 1);
  return {
    horizontal: summarizeGaps(gapsForGroups(rowGroups, "horizontal")),
    vertical: summarizeGaps(gapsForGroups(columnGroups, "vertical"))
  };
}

function setHidden(element, hidden) {
  element?.classList.toggle("hidden", hidden);
}

function populateSpacing(shapes) {
  const spacing = spacingInfoForShapes(shapes);
  latestSpacing = spacing;
  const showHorizontal = spacing.horizontal !== null;
  const showVertical = spacing.vertical !== null;

  setHidden($("spacingGroup"), !showHorizontal && !showVertical);
  setHidden($("horizontalSpacingField"), !showHorizontal);
  setHidden($("verticalSpacingField"), !showVertical);
  fields.horizontalSpacing.value = showHorizontal ? spacing.horizontal : "";
  fields.verticalSpacing.value = showVertical ? spacing.vertical : "";
}

function bindColorPair(colorInput, hexInput) {
  colorInput.addEventListener("input", () => {
    hexInput.value = cleanHex(colorInput.value);
  });
  hexInput.addEventListener("input", () => {
    const hex = cleanHex(hexInput.value, cleanHex(colorInput.value));
    if (hex.length === 6) colorInput.value = hexToColorInput(hex);
  });
}

function populate(shape) {
  selectedShapeId = shape.id;
  selectedShapeIds = shape.ids || [shape.id].filter(Boolean);
  aspectRatio = shape.width && shape.height ? shape.width / shape.height : 1;
  currentBounds = {
    left: Math.round(shape.left || 0),
    top: Math.round(shape.top || 0),
    width: Math.max(1, Math.round(shape.width || 1)),
    height: Math.max(1, Math.round(shape.height || 1))
  };

  fields.left.value = currentBounds.left;
  fields.top.value = currentBounds.top;
  fields.width.value = currentBounds.width;
  fields.height.value = currentBounds.height;
  fields.opacity.value = Math.round(shape.opacity ?? 100);
  fields.fillOpacity.value = Math.round(shape.fillOpacity ?? 100);
  fields.lineOpacity.value = Math.round(shape.lineOpacity ?? 100);
  if (fields.lineWeight) fields.lineWeight.value = Number(shape.lineWeight ?? 1);

  syncColorPair(fields.fillColor, fields.fillHex, shape.fillColor || "FFFFFF");
  syncColorPair(fields.lineColor, fields.lineHex, shape.lineColor || "0D0D0D");
}

function populateFromShapes(shapes) {
  const bounds = boundsForShapes(shapes);
  const first = shapes[0];
  populate({
    id: first.id,
    ids: shapes.map((shape) => shape.id),
    name: shapes.length === 1 ? first.name : `${shapes.length}個の図形を選択中`,
    left: bounds.left,
    top: bounds.top,
    width: bounds.width,
    height: bounds.height,
    visible: shapes.every((shape) => shape.visible !== false),
    opacity: 100,
    fillColor: first.fillColor || "FFFFFF",
    fillOpacity: Math.round((1 - (first.fillTransparency || 0)) * 100),
    lineColor: first.lineColor || "0D0D0D",
    lineOpacity: Math.round((1 - (first.lineTransparency || 0)) * 100),
    lineWeight: first.lineWeight || 1
  });
  if (shapes.length > 1) {
    fields.left.value = commonRoundedOrMixed(shapes, "left");
    fields.top.value = commonRoundedOrMixed(shapes, "top");
    fields.width.value = commonRoundedOrMixed(shapes, "width");
    fields.height.value = commonRoundedOrMixed(shapes, "height");
  }
  populateSpacing(shapes);
}

function valuesFromForm() {
  return {
    left: parseNumberField(fields.left, currentBounds.left),
    top: parseNumberField(fields.top, currentBounds.top),
    width: parseNumberField(fields.width, currentBounds.width, 1),
    height: parseNumberField(fields.height, currentBounds.height, 1),
    opacity: clamp(fields.opacity.value, 0, 100),
    fillColor: cleanHex(fields.fillHex.value, "FFFFFF"),
    fillOpacity: clamp(fields.fillOpacity.value, 0, 100),
    lineColor: cleanHex(fields.lineHex.value, "0D0D0D"),
    lineOpacity: clamp(fields.lineOpacity.value, 0, 100),
    lineWeight: Math.max(0, Number(fields.lineWeight?.value || 0))
  };
}

async function refreshSelection(options = {}) {
  const { silent = false } = options;
  if (isApplying || isRefreshing || (silent && isEditingFormField())) return;
  if (!isOfficeReady()) {
    populate(demoShape);
    if (!silent) setStatus("ブラウザ単体のデモ表示です。PowerPointで開くと選択図形を編集できます。");
    return;
  }

  try {
    isRefreshing = true;
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (!selectedShapes.items.length) {
        populateSpacing([]);
        if (!silent) {
          if (!selectedShapeIds.length) selectedShapeId = null;
          setStatus("スライド上の図形を選択してから再読み込みしてください。");
        }
        return;
      }

      selectedShapes.items.forEach((shape) => {
        shape.load("id,name,left,top,width,height,visible");
        shape.fill.load("foregroundColor,transparency");
        shape.lineFormat.load("color,transparency,weight");
      });
      await context.sync();

      populateFromShapes(
        selectedShapes.items.map((shape) => ({
          id: shape.id,
          name: shape.name,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
          visible: shape.visible,
          fillColor: shape.fill.foregroundColor || "FFFFFF",
          fillTransparency: shape.fill.transparency || 0,
          lineColor: shape.lineFormat.color || "0D0D0D",
          lineTransparency: shape.lineFormat.transparency || 0,
          lineWeight: shape.lineFormat.weight || 1
        }))
      );

      if (!silent) setStatus(`${selectedShapes.items.length}個の図形を読み込みました。`);
    });
  } catch (error) {
    if (!silent) setStatus(`読み込みに失敗しました: ${formatError(error)}`);
  } finally {
    isRefreshing = false;
  }
}

async function resolveTargetShapes(context) {
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedShapes.load("items");
  await context.sync();

  if (selectedShapes.items.length) {
    selectedShapes.items.forEach((shape) => shape.load("id,name,left,top,width,height,visible"));
    await context.sync();
    selectedShapeIds = selectedShapes.items.map((shape) => shape.id);
    selectedShapeId = selectedShapeIds[0] || null;
    return selectedShapes.items;
  }

  if (!selectedShapeIds.length) return [];

  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items");
  await context.sync();
  if (!selectedSlides.items.length) return [];

  const shapes = selectedShapeIds.map((id) => selectedSlides.items[0].shapes.getItem(id));
  shapes.forEach((shape) => shape.load("id,name,left,top,width,height,visible"));
  await context.sync();
  return shapes;
}

function changedSet(fieldsToApply) {
  return new Set(fieldsToApply && fieldsToApply.length ? fieldsToApply : [
    "left",
    "top",
    "width",
    "height",
    "opacity",
    "fillColor",
    "fillOpacity",
    "lineColor",
    "lineOpacity",
    "lineWeight"
  ]);
}

function applyDemoChanges(next, changes) {
  if (changes.has("left")) demoShape.left = next.left;
  if (changes.has("top")) demoShape.top = next.top;
  if (changes.has("width")) demoShape.width = next.width;
  if (changes.has("height")) demoShape.height = next.height;
  if (changes.has("opacity")) demoShape.fillOpacity = next.opacity;
  if (changes.has("fillColor")) demoShape.fillColor = next.fillColor;
  if (changes.has("fillOpacity")) demoShape.fillOpacity = next.fillOpacity;
  if (changes.has("lineColor")) demoShape.lineColor = next.lineColor;
  if (changes.has("lineOpacity")) demoShape.lineOpacity = next.lineOpacity;
  if (changes.has("lineWeight")) demoShape.lineWeight = next.lineWeight;
}

function applyShapeChanges(shape, next, changes) {
  if (changes.has("left")) shape.left = next.left;
  if (changes.has("top")) shape.top = next.top;
  if (changes.has("width")) shape.width = next.width;
  if (changes.has("height")) shape.height = next.height;
  if (changes.has("fillColor")) shape.fill.setSolidColor(hexToOfficeColor(next.fillColor));
  if (changes.has("fillOpacity") || changes.has("opacity")) {
    shape.fill.transparency = 1 - (next.fillOpacity * next.opacity) / 10000;
  }
  if (changes.has("lineColor")) shape.lineFormat.color = hexToOfficeColor(next.lineColor);
  if (changes.has("lineOpacity") || changes.has("opacity")) {
    shape.lineFormat.transparency = 1 - (next.lineOpacity * next.opacity) / 10000;
  }
  if (changes.has("lineWeight") && fields.lineWeight) shape.lineFormat.weight = next.lineWeight;
}

async function applyToSelection(fieldsToApply) {
  const changes = changedSet(fieldsToApply);
  const next = valuesFromForm();

  if (!isOfficeReady()) {
    applyDemoChanges(next, changes);
    populate(demoShape);
    setStatus("デモ図形へ反映しました。PowerPointで開くと実際の図形へ反映します。");
    return;
  }

  try {
    isApplying = true;
    await PowerPoint.run(async (context) => {
      const shapes = await resolveTargetShapes(context);
      if (!shapes.length) {
        setStatus("適用する図形がありません。");
        return;
      }

      shapes.forEach((shape) => {
        shape.visible = true;
        applyShapeChanges(shape, next, changes);
      });

      await context.sync();
      setStatus(`${shapes.length}個の図形へリアルタイム反映しました。`);
    });
  } catch (error) {
    setStatus(`適用に失敗しました: ${formatError(error)}`);
  } finally {
    isApplying = false;
  }
}

function scheduleRefreshSelection() {
  window.clearTimeout(refreshTimer);
  refreshTimer = window.setTimeout(() => refreshSelection(), 180);
}

function startLiveSelectionWatch() {
  window.clearInterval(liveRefreshTimer);
  liveRefreshTimer = window.setInterval(() => {
    if (document.visibilityState === "visible") refreshSelection({ silent: true });
  }, 650);
}

function scheduleApply(fieldsToApply) {
  const names = Array.isArray(fieldsToApply) ? fieldsToApply : [fieldsToApply].filter(Boolean);
  names.forEach((name) => pendingApplyFields.add(name));
  window.clearTimeout(applyTimer);
  applyTimer = window.setTimeout(() => {
    const namesToApply = [...pendingApplyFields];
    pendingApplyFields.clear();
    applyToSelection(namesToApply);
  }, 220);
}

function scheduleSpacingApply(direction) {
  pendingApplyFields.clear();
  window.clearTimeout(applyTimer);
  applyTimer = window.setTimeout(() => applySpacing(direction), 220);
}

async function align(command) {
  if (!isOfficeReady()) {
    const slideWidth = 960;
    const slideHeight = 540;
    const width = Number(fields.width.value || 0);
    const height = Number(fields.height.value || 0);
    if (command === "alignLeft") fields.left.value = 0;
    if (command === "alignCenter") fields.left.value = Math.round((slideWidth - width) / 2);
    if (command === "alignRight") fields.left.value = Math.round(slideWidth - width);
    if (command === "alignTop") fields.top.value = 0;
    if (command === "alignMiddle") fields.top.value = Math.round((slideHeight - height) / 2);
    if (command === "alignBottom") fields.top.value = Math.round(slideHeight - height);
    applyToSelection(["left", "top"]);
    return;
  }

  try {
    isApplying = true;
    await PowerPoint.run(async (context) => {
      const shapes = await resolveTargetShapes(context);
      if (!shapes.length) {
        setStatus("整列する図形がありません。");
        return;
      }

      const bounds = boundsForShapes(shapes);
      const singleSlideWidth = 960;
      const singleSlideHeight = 540;
      shapes.forEach((shape) => {
        if (shapes.length === 1) {
          if (command === "alignLeft") shape.left = 0;
          if (command === "alignCenter") shape.left = (singleSlideWidth - shape.width) / 2;
          if (command === "alignRight") shape.left = singleSlideWidth - shape.width;
          if (command === "alignTop") shape.top = 0;
          if (command === "alignMiddle") shape.top = (singleSlideHeight - shape.height) / 2;
          if (command === "alignBottom") shape.top = singleSlideHeight - shape.height;
          return;
        }

        if (command === "alignLeft") shape.left = bounds.left;
        if (command === "alignCenter") shape.left = bounds.left + (bounds.width - shape.width) / 2;
        if (command === "alignRight") shape.left = bounds.right - shape.width;
        if (command === "alignTop") shape.top = bounds.top;
        if (command === "alignMiddle") shape.top = bounds.top + (bounds.height - shape.height) / 2;
        if (command === "alignBottom") shape.top = bounds.bottom - shape.height;
      });

      await context.sync();
      setStatus(`${shapes.length}個の図形を整列しました。`);
      fields.left.value = Math.round(bounds.left);
      fields.top.value = Math.round(bounds.top);
      fields.width.value = Math.round(bounds.width);
      fields.height.value = Math.round(bounds.height);
    });
  } catch (error) {
    setStatus(`整列に失敗しました: ${formatError(error)}`);
  } finally {
    isApplying = false;
  }
}

async function applySpacing(direction) {
  const field = direction === "horizontal" ? fields.horizontalSpacing : fields.verticalSpacing;
  const spacing = Number(String(field.value).replace(/[^\d.-]/g, ""));
  if (!Number.isFinite(spacing)) {
    field.value = "混在";
    return;
  }

  if (!isOfficeReady()) {
    setStatus("PowerPointで開くと間隔を反映できます。");
    return;
  }

  try {
    isApplying = true;
    await PowerPoint.run(async (context) => {
      const shapes = await resolveTargetShapes(context);
      if (shapes.length < 2) {
        setStatus("間隔を設定するには複数図形を選択してください。");
        return;
      }

      const groups =
        direction === "horizontal"
          ? groupByAxis(shapes, "row").filter((group) => group.length > 1)
          : groupByAxis(shapes, "column").filter((group) => group.length > 1);

      groups.forEach((group) => {
        const sorted = [...group].sort((a, b) => (direction === "horizontal" ? a.left - b.left : a.top - b.top));
        sorted.slice(1).forEach((shape, index) => {
          const previous = sorted[index];
          if (direction === "horizontal") {
            shape.left = previous.left + previous.width + spacing;
          } else {
            shape.top = previous.top + previous.height + spacing;
          }
        });
      });

      await context.sync();
      setStatus(`${direction === "horizontal" ? "横" : "縦"}の間隔を${Math.round(spacing)}に揃えました。`);
    });
  } catch (error) {
    setStatus(`間隔の反映に失敗しました: ${formatError(error)}`);
  } finally {
    isApplying = false;
  }
  await refreshSelection({ silent: true });
}

function bindEvents() {
  bindColorPair(fields.fillColor, fields.fillHex);
  bindColorPair(fields.lineColor, fields.lineHex);
  fields.fillColor.addEventListener("input", () => scheduleApply("fillColor"));
  fields.lineColor.addEventListener("input", () => scheduleApply("lineColor"));

  bindSelectableInput(fields.horizontalSpacing, {
    onValid: () => scheduleSpacingApply("horizontal")
  });
  bindSelectableInput(fields.verticalSpacing, {
    onValid: () => scheduleSpacingApply("vertical")
  });

  bindSelectableInput(fields.width, {
    min: 1,
    onValid: () => {
      if (ratioLocked) {
        fields.height.value = Math.round(Number(fields.width.value || 1) / aspectRatio);
        scheduleApply(["width", "height"]);
        return;
      }
      scheduleApply("width");
    }
  });

  bindSelectableInput(fields.height, {
    min: 1,
    onValid: () => {
      if (ratioLocked) {
        fields.width.value = Math.round(Number(fields.height.value || 1) * aspectRatio);
        scheduleApply(["width", "height"]);
        return;
      }
      scheduleApply("height");
    }
  });

  [
    [fields.left, "left"],
    [fields.top, "top"],
    [fields.opacity, "opacity"],
    [fields.fillHex, "fillColor"],
    [fields.fillOpacity, "fillOpacity"],
    [fields.lineHex, "lineColor"],
    [fields.lineOpacity, "lineOpacity"],
    [fields.lineWeight, "lineWeight"]
  ].filter(([field]) => Boolean(field)).forEach(([field, name]) => {
    const isHexField = field === fields.fillHex || field === fields.lineHex;
    const isPercentField = field === fields.opacity || field === fields.fillOpacity || field === fields.lineOpacity;
    bindSelectableInput(field, {
      allowStep: !isHexField,
      min: name === "lineWeight" ? 0 : undefined,
      max: isPercentField ? 100 : undefined,
      validate: isHexField ? (value) => /^[0-9a-f]{6}$/i.test(String(value).trim()) : isFiniteNumberText,
      onValid: () => scheduleApply(name)
    });
  });

  document.querySelectorAll("[data-command]").forEach((button) => {
    button.addEventListener("click", () => {
      const command = button.dataset.command;
      if (command?.startsWith("align")) align(command);
    });
  });

}

bindEvents();

if (window.Office) {
  Office.onReady(() => {
    refreshSelection();
    startLiveSelectionWatch();
    if (Office.context?.document?.addHandlerAsync) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRefreshSelection);
    }
  });
} else {
  refreshSelection();
}
