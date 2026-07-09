// Scope constants used by UI and logic
// Only two scopes are supported now:
//  - Selected: only currently selected shapes
//  - Current: all shapes on the active slide
const Scope = {
    Selected: "selected",
    Current: "current"
};

// Stores information about the last normalization operation so it can be undone.
let lastNormalizationUndoData = null;

/**
 * Core function that applies (or simulates) the border-radius normalization
 * based on the given scope and radius in points.
 *
 * Uses the candidate logic that worked in debugApplyCornerRadiusOnCandidates:
 *  - geometric shapes only
 *  - exactly 1 adjustment
 *  - adjustment(0) <= 0.5  => candidate
 *
 * @param {number} radiusPts
 * @param {{ applyChanges: boolean, scope: string }} options
 * @returns {Promise<object>} summary with counts
 */

async function normalizeRectangleCorners(radiusPts, options) {
    const applyChanges = options.applyChanges;
    const scope = options.scope;

    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        showNotification(
            "Not supported",
            "This version of PowerPoint does not support the required shapes API (PowerPointApi 1.10)."
        );
        return {
            slidesProcessed: 0,
            shapesProcessed: 0,
            shapesModified: 0,
            shapesSkipped: 0,
            shapesWithAdjustment: 0
        };
    }

    const summary = {
        slidesProcessed: 0,
        shapesProcessed: 0,
        shapesModified: 0,
        shapesSkipped: 0,
        shapesWithAdjustment: 0,
        aborted: false  // <--- NEU
    };

    // Initialize undo buffer for this run
    lastNormalizationUndoData = {
        entries: [],
        applyChanges,
        radiusPts,
        scope
    };


    await PowerPoint.run(async (context) => {
        const presentation = context.presentation;

        // Deduplicate processing across the whole run (selected or current).
        const processedIds = new Set();

        // ============================
        // Scope: SELECTED SHAPES ONLY
        // ============================
        if (scope === Scope.Selected) {
            const selectedShapes = presentation.getSelectedShapes();
            selectedShapes.load("items");
            await context.sync();

            const shapes = selectedShapes.items || [];
            if (!shapes.length) {
                showNotification(
                    "No shapes selected",
                    "No shapes are selected. Please select one or more shapes or choose 'Active slide'."
                );
                summary.aborted = true;
                return;
            }

            summary.slidesProcessed = 1;

            for (const shape of shapes) {
                await processShapeCandidateBased(context, shape, radiusPts, applyChanges, summary, processedIds);
            }

            return;
        }

        // ============================
        // Scope: ACTIVE SLIDE (CURRENT)
        // ============================
        const slideRefs = await getSlidesByScope(context, presentation, scope);
        if (!slideRefs || slideRefs.length === 0) {
            showNotification(
                "No slide in scope",
                "No active slide was found. Please click onto a slide and try again."
            );
            summary.aborted = true;
            return;
        }

        for (const slide of slideRefs) {
            summary.slidesProcessed++;
            await applyCornerRadiusOnCandidatesOnSlide(context, slide, radiusPts, applyChanges, summary, processedIds);
        }
    });

    return summary;
}

/**
 * Returns the slide that should be processed, based on the selected scope.
 * With the current UI we only support "current" (active slide).
 */
async function getSlidesByScope(context, presentation, scope) {
    // Currently only used for "Current" – for other values we behave the same way.
    const selectedSlides = presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (!selectedSlides.items || selectedSlides.items.length === 0) {
        return [];
    }

    // First (visible) slide only – the active slide.
    return [selectedSlides.items[0]];
}

/**
 * Applies the candidate-based corner-radius logic to all shapes on a slide.
 * Uses the same candidate criteria as debugApplyCornerRadiusOnCandidates.
 */
async function applyCornerRadiusOnCandidatesOnSlide(context, slide, radiusPts, applyChanges, summary, processedIds) {
    slide.shapes.load("items");
    await context.sync();

    const shapes = slide.shapes.items || [];
    if (!shapes.length) return;

    for (const shape of shapes) {
        await processShapeCandidateBased(context, shape, radiusPts, applyChanges, summary, processedIds);
    }
}

/**
 * Processes a shape using the rounded-rectangle candidate heuristic.
 *
 * A shape is considered a valid candidate if:
 *  - it has exactly one adjustment
 *  - adjustment(0) is a number in the range 0..0.5
 *  - it has either a fill OR an outline (line)
 *
 * The shape type is intentionally NOT used as a filter:
 * TextBox shapes can still represent rounded rectangles
 * if they expose a valid corner-radius adjustment.
 */
async function processShapeCandidateBased(
    context,
    shape,
    radiusPts,
    applyChanges,
    summary,
    processedIds
) {
    // Load only what we need for early decisions
    shape.load([
        "id",
        "name",
        "type",
        "level",
        "left",
        "top",
        "width",
        "height",
        "rotation",
        "adjustments",
        "fill/type",
        "lineFormat/visible"
    ]);
    await context.sync();

    // ------------------------------------------------------------
    // Deduplication: ensure each shape is processed only once per run
    // ------------------------------------------------------------
    if (processedIds && processedIds.has(shape.id)) {
        return;
    }

    if (processedIds) {
        processedIds.add(shape.id);
    }

    summary.shapesProcessed++;

    // ------------------------------------------------------------
    // Group handling: recurse into children, do not process the group itself
    // ------------------------------------------------------------
 
    if (shape.type === PowerPoint.ShapeType.group) {
        try {
             const groupShapes = shape.group.shapes;
            groupShapes.load("items");
            await context.sync();

            for (const child of groupShapes.items || []) {
                 await processShapeCandidateBased(
                    context,
                    child,
                    radiusPts,
                    applyChanges,
                    summary,
                    processedIds
                );
            }
        } catch (e) {
            console.warn("Could not process group shapes:",
                {
                    shapeId: shape.id,
                    shapeName: shape.name,
                    error: e
                });
            summary.shapesSkipped++;
        }

        return;
    }

    // ------------------------------------------------------------
    // Safety check: shape must have a visible fill or outline
    // (prevents pure text boxes without geometry)
    // ------------------------------------------------------------
    const hasFill = shape.fill && shape.fill.type !== PowerPoint.ShapeFillType.noFill;
    const hasLine = shape.lineFormat && shape.lineFormat.visible === true;

    if (!hasFill && !hasLine) {

        summary.shapesSkipped++;
        return;
    }

    // ------------------------------------------------------------
    // Adjustment-based candidate detection (core logic)
    // ------------------------------------------------------------
    const adjustments = shape.adjustments;

    if (!adjustments) {

        summary.shapesSkipped++;
        return;
    }

    adjustments.load("count");
    await context.sync();

    if (adjustments.count !== 1) {
        
        summary.shapesSkipped++;
        return;
    }

    // Read the actual adjustment value (corner radius)
    const adj0 = adjustments.get(0);
    await context.sync();

    const firstVal = adj0.value;

    // Candidate heuristic: looks like a rounded rectangle
    if (typeof firstVal !== "number" || firstVal < 0 || firstVal > 0.5) {

        summary.shapesSkipped++;
        return;
    }

    // ------------------------------------------------------------
    // Valid candidate: apply the uniform rounded corner logic
    // ------------------------------------------------------------
    await applyUniformRoundedCorner(
        context,
        shape,
        radiusPts,
        applyChanges,
        summary
    );
}


    

/**
 * Applies the corner radius to a geometric shape using Adjustments(0).
 */
async function applyUniformRoundedCorner(context, shape, radiusPts, applyChanges, summary) {
    const width = shape.width;
    const height = shape.height;

    if (width <= 0 || height <= 0) {
        summary.shapesSkipped++;
        return;
    }

    const minDim = Math.min(width, height);
    let adjValue = radiusPts / minDim;

    // Clamp to 0..0.5
    if (adjValue < 0) adjValue = 0;
    if (adjValue > 0.5) adjValue = 0.5;

    const adjustments = shape.adjustments;

    try {
        // Read current value (needed for undo).
        const adj0 = adjustments.get(0);
        await context.sync();
        const oldValue = adj0.value;

        summary.shapesWithAdjustment++;

        if (applyChanges && lastNormalizationUndoData) {
            lastNormalizationUndoData.entries.push({ shapeId: shape.id, oldValue });
        }

        if (!applyChanges) {
            // Dry run: no changes.
            return;
        }

        const isWeb = Office.context.platform === Office.PlatformType.OfficeOnline;

        // 1) Insert dot if needed (forces repaint in PowerPoint for the web).
        //    This call does its own sync only when it actually injects.
        const injectedState = isWeb ? await injectRenderDotIfNeeded(context, shape) : null;

        // 2) Queue the radius change.
        adjustments.set(0, adjValue);

        // 2a) minimal nudge to prevent weird repositioning after closing and reopening
        const originalLeft = shape.left;

        shape.left = originalLeft + 0.01;
        await context.sync();

        //shape.left = originalLeft;

        //await context.sync();

        // 3) Queue dot removal + state restore (no sync here).
        if (isWeb && injectedState) {
            queueRestoreInjectedText(shape, injectedState);
        }

        // Commit (radius + restore) in one round-trip.
        await context.sync();

        summary.shapesModified++;
    } catch (e) {
        console.log("Shape skipped, no valid Adjustment[0]:", e);
        summary.shapesSkipped++;
    }
}

async function injectRenderDotIfNeeded(context, shape) {
    try {
        // Load everything needed in one go.
        shape.load(["width", "height", "textFrame/hasText", "textFrame/autoSizeSetting", "textFrame/textRange/text"]);
        await context.sync();

        const tf = shape.textFrame;
        if (!tf) return null;

        // If the shape already has text, do nothing.
        if (tf.hasText === true) return null;

        const state = {
            originalText: tf.textRange?.text || "",
            originalAuto: tf.autoSizeSetting,
            originalW: shape.width,
            originalH: shape.height
        };

        // Prevent auto-resize side effects.
        tf.autoSizeSetting = "AutoSizeNone";
        tf.textRange.text = "."; // Visible glyph -> triggers repaint in PowerPoint Web
        await context.sync();

        // Restore geometry immediately (still in Web-only injection step).
        // This avoids visible jitter while the dot exists.
        shape.width = state.originalW;
        shape.height = state.originalH;

        // Do not sync here; caller will sync soon anyway.
        // But we already synced to commit the dot. We keep geometry changes queued.
        return state;
    } catch (e) {
        console.warn("injectRenderDotIfNeeded skipped:", e);
        return null;
    }
}

function queueRestoreInjectedText(shape, state) {
    // Restore original text and autosize, and reset geometry.
    // This function does NOT call context.sync(); caller batches it with other operations.
    const tf = shape.textFrame;
    if (!tf) return;

    tf.textRange.text = state.originalText;

    if (state.originalAuto !== undefined && state.originalAuto !== null) {
        tf.autoSizeSetting = state.originalAuto;
    }

    shape.width = state.originalW;
    shape.height = state.originalH;
}


async function kickWebRenderOnSelectedSlide() {
    const isWeb = Office.context.platform === Office.PlatformType.OfficeOnline;
    if (!isWeb) return;

    await PowerPoint.run(async (context) => {
        const pres = context.presentation;

        const sel = pres.getSelectedSlides();
        sel.load("items");
        await context.sync();

        if (!sel.items || sel.items.length === 0) return;

        const slide = sel.items[0];

        // Render-Kicker: Folie als kleines Bild rendern lassen
        // Slide.getImageAsBase64 ist in der Slide API dokumentiert. [1](https://stackoverflow.com/questions/65688141/is-it-possible-to-select-update-shape-icons-image-properties-using-ms-powerpoint)
        //const img = slide.getImageAsBase64({ width: 32 });
        //await context.sync();
        //void img.value; // Ergebnis ignorieren
    });
}





/**
 * Undo the last normalization operation by restoring the previous adjustment values.
 */
async function undoLastNormalization() {
    if (!lastNormalizationUndoData || !lastNormalizationUndoData.entries.length) {
        showNotification(
            "Nothing to undo",
            "There is no previous normalization operation to undo."
        );
        return;
    }

    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        showNotification(
            "Not supported",
            "This version of PowerPoint does not support the required shapes API (PowerPointApi 1.10)."
        );
        return;
    }

    const entries = lastNormalizationUndoData.entries;

    await PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const slides = presentation.slides;
        slides.load("items");
        await context.sync();

        // Map shapeId -> oldValue for quick lookup
        const byId = {};
        entries.forEach(e => {
            byId[e.shapeId] = e.oldValue;
        });

        for (const slide of slides.items) {
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();

            for (const shape of shapes.items) {
                await restoreShapeAdjustmentRecursive(context, shape, byId);
            }
        }

        await context.sync();
    });

    showNotification(
        "Undo completed",
        "The last normalization operation was undone."
    );

    // Clear undo buffer so the operation cannot be undone twice.
    lastNormalizationUndoData = null;
}

/**
 * Recursively restores adjustment[0] for shapes whose IDs are in the lookup map.
 */
async function restoreShapeAdjustmentRecursive(context, shape, idToOldValueMap) {
    shape.load(["id", "type", "adjustments"]);
    await context.sync();

    // Recurse into groups
    if (shape.type === PowerPoint.ShapeType.group) {
        try {
            const groupShapes = shape.group.shapes;
            groupShapes.load("items");
            await context.sync();

            for (const subShape of groupShapes.items || []) {
                await restoreShapeAdjustmentRecursive(
                    context,
                    subShape,
                    idToOldValueMap
                );
            }
        } catch (e) {
            console.warn("Could not restore group children:", shape.id, e);
        }

        return;
    }

    // Restore if this shape was changed during the last normalization
    if (Object.prototype.hasOwnProperty.call(idToOldValueMap, shape.id)) {
        const adjustments = shape.adjustments;
        if (adjustments) {
            try {
                adjustments.set(0, idToOldValueMap[shape.id]);
            } catch (e) {
                console.log("Failed to restore adjustment for shape", shape.id, e);
            }
        }
    }
}