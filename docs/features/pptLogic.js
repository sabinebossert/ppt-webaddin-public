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
                summary.aborted = true;   // <--- NEU
                return;
            }

            summary.slidesProcessed = 1;

            for (const shape of shapes) {
                await processShapeCandidateBased(context, shape, radiusPts, applyChanges, summary);
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
            summary.aborted = true;   // <--- NEU
            return;
        }

        for (const slide of slideRefs) {
            summary.slidesProcessed++;
            await applyCornerRadiusOnCandidatesOnSlide(context, slide, radiusPts, applyChanges, summary);
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
async function applyCornerRadiusOnCandidatesOnSlide(context, slide, radiusPts, applyChanges, summary) {
    slide.shapes.load("items");
    await context.sync();

    const shapes = slide.shapes.items || [];
    if (!shapes.length) {
        return;
    }

    for (const shape of shapes) {
        await processShapeCandidateBased(context, shape, radiusPts, applyChanges, summary);
    }
}

/**
 * Processes a shape using the "candidate" heuristic:
 *  - handles groups recursively
 *  - only geometric shapes
 *  - exactly 1 adjustment
 *  - adjustment(0) <= 0.5
 */
async function processShapeCandidateBased(context, shape, radiusPts, applyChanges, summary) {
    shape.load(["id", "name", "type", "width", "height", "groupItems", "adjustments"]);
    await context.sync();

    // 1. Handle groups recursively
    if (shape.type === PowerPoint.ShapeType.group && shape.groupItems) {
        const groupItems = shape.groupItems;
        groupItems.load("items");
        await context.sync();

        if (groupItems.items) {
            for (const subShape of groupItems.items) {
                await processShapeCandidateBased(context, subShape, radiusPts, applyChanges, summary);
            }
        }
        return;
    }

    summary.shapesProcessed++;

    // 2. Only geometric shapes
    if (shape.type !== PowerPoint.ShapeType.geometricShape) {
        summary.shapesSkipped++;
        return;
    }

    const adjustments = shape.adjustments;
    let count = 0;
    let firstVal = undefined;

    if (adjustments) {
        adjustments.load("count");
        await context.sync();
        count = adjustments.count;

        if (count > 0) {
            const res = adjustments.get(0);
            await context.sync();
            firstVal = res.value;
        }
    }

    // Candidate logic as in debugApplyCornerRadiusOnCandidates:
    const isCandidate =
        shape.type === PowerPoint.ShapeType.geometricShape &&
        count === 1 &&
        typeof firstVal === "number" &&
        firstVal <= 0.5;

    if (!isCandidate) {
        summary.shapesSkipped++;
        return;
    }

    // Apply the corner radius
    await applyUniformRoundedCorner(context, shape, radiusPts, applyChanges, summary);
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

    // clamp 0..0.5 (similar to the VBA logic)
    if (adjValue < 0) adjValue = 0;
    if (adjValue > 0.5) adjValue = 0.5;

    const adjustments = shape.adjustments;

    try {
        // Read current value
        const adj0 = adjustments.get(0);  // ClientResult<number>
        await context.sync();
        const oldValue = adj0.value;

        summary.shapesWithAdjustment++;

        // Remember old value for undo (only if we are actually changing things)
        if (applyChanges && lastNormalizationUndoData) {
            lastNormalizationUndoData.entries.push({
                shapeId: shape.id,
                oldValue
            });
        }

        if (!applyChanges) {
            console.log("Dry run – current radius:", oldValue, "new (simulated):", adjValue);
            return;
        }

        // Set new value
        adjustments.set(0, adjValue);
        await context.sync();

        summary.shapesModified++;
        console.log("Corner radius changed from", oldValue, "to", adjValue);
    } catch (e) {
        console.log("Shape skipped, no valid Adjustment[0]:", e);
        summary.shapesSkipped++;
    }
}

// -----------------------------
// Debug and test functions
// -----------------------------

async function debugShape(context, shape) {
    // 1. Load basic shape data
    shape.load([
        "id",
        "name",
        "type",
        "left",
        "top",
        "width",
        "height",
        "title",
        "description",
        "tags",
        "adjustments"
    ]);

    await context.sync();

    console.log("=== DebugShape start ===");
    console.log("ID:", shape.id);
    console.log("Name:", shape.name);
    console.log("Type (ShapeType):", shape.type);
    console.log("Position:", { left: shape.left, top: shape.top });
    console.log("Size:", { width: shape.width, height: shape.height });
    console.log("Title:", shape.title);
    console.log("Description:", shape.description);
    console.log("Tags:", shape.tags);

    const adjustments = shape.adjustments;
    if (adjustments) {
        // 2. Load number of adjustments
        adjustments.load("count");
        await context.sync();

        console.log("Adjustments.count:", adjustments.count);

        if (adjustments.count > 0) {
            // 3. Get all ClientResult objects
            const results = [];
            for (let i = 0; i < adjustments.count; i++) {
                const res = adjustments.get(i);
                results.push(res);
            }

            // 4. Sync to retrieve values
            await context.sync();

            // 5. Log values
            results.forEach((res, index) => {
                console.log(`Adjustments[${index}].value:`, res.value);
            });
        }
    } else {
        console.log("Adjustments: (none)");
    }

    console.log("=== DebugShape end ===");
}

// Debug: shows how the rounded-corner heuristic would decide for a shape.
async function debugRoundedCornerHeuristic(context, shape) {
    shape.load(["id", "name", "type", "adjustments"]);
    await context.sync();

    console.log("=== DebugRoundedCornerHeuristic start ===");
    console.log("ID:", shape.id);
    console.log("Name:", shape.name);
    console.log("Type (ShapeType):", shape.type);

    const adjustments = shape.adjustments;
    if (!adjustments) {
        console.log("Adjustments: (none) -> would be skipped");
        console.log("=== DebugRoundedCornerHeuristic end ===");
        return;
    }

    // Load count of adjustments
    adjustments.load("count");
    await context.sync();

    console.log("Adjustments.count:", adjustments.count);

    // Non-geometric shapes would be skipped anyway
    if (shape.type !== PowerPoint.ShapeType.geometricShape) {
        console.log("=> Not a GeometricShape -> would be skipped");
        console.log("=== DebugRoundedCornerHeuristic end ===");
        return;
    }

    // Only shapes with exactly one adjustment are candidates in this logic
    if (adjustments.count !== 1) {
        console.log("=> Adjustments.count != 1 -> would be skipped");
        console.log("=== DebugRoundedCornerHeuristic end ===");
        return;
    }

    // Original value of Adjustment(0)
    const originalResult = adjustments.get(0);
    await context.sync();
    const originalValue = originalResult.value;

    console.log("Original Adjustments[0].value:", originalValue);

    const probeValue = 0.9; // deliberately > 0.5
    let probedValue = null;
    let isRoundedLike = false;

    try {
        // Probe: set a high value
        adjustments.set(0, probeValue);
        await context.sync();

        // Read probed value
        const probedResult = adjustments.get(0);
        await context.sync();
        probedValue = probedResult.value;

        console.log("Probed Adjustments[0].value (after set 0.9):", probedValue);

        shape.load("adjustments");
        await context.sync();

        const resAgain = shape.adjustments.get(0);
        await context.sync();
        console.log("Probed Adjustments[0].value after reload:", resAgain.value);

        const MAX_ROUNDED = 0.5;
        const EPS = 0.0001;

        if (probedValue <= MAX_ROUNDED + EPS) {
            isRoundedLike = true;
        } else {
            isRoundedLike = false;
        }
    } catch (e) {
        console.log("Error while probing adjustment:", e);
        isRoundedLike = true;
    } finally {
        // Restore original value
        try {
            adjustments.set(0, originalValue);
            await context.sync();
            console.log("Original Adjustments[0].value restored:", originalValue);
        } catch (restoreError) {
            console.log("Error while restoring original adjustment value:", restoreError);
        }
    }

    console.log("Heuristic isRoundedLike:", isRoundedLike);
    console.log("=> This shape would",
        isRoundedLike ? "BE PROCESSED" : "BE SKIPPED",
        "by normalizeRectangleCorners().");
    console.log("=== DebugRoundedCornerHeuristic end ===");
}

// Debug: deletes all shapes that are very likely NOT rounded / semi-rounded rectangles.
// Only candidate shapes remain.
async function debugFilterAndDeleteNonRoundedShapes(context, shape) {
    shape.load(["id", "name", "type", "adjustments"]);
    await context.sync();

    const adjustments = shape.adjustments;
    let count = 0;
    let values = [];

    if (adjustments) {
        adjustments.load("count");
        await context.sync();

        count = adjustments.count;

        if (count > 0) {
            const results = [];
            for (let i = 0; i < count; i++) {
                const res = adjustments.get(i);
                results.push(res);
            }

            await context.sync();
            values = results.map(r => r.value);
        }
    }

    const isGeometric = (shape.type === PowerPoint.ShapeType.geometricShape);
    const firstVal = values.length > 0 ? values[0] : undefined;

    let isSafeNonRounded =
        !isGeometric ||
        count === 0 ||
        count > 1 ||
        (count === 1 && typeof firstVal === "number" && firstVal > 0.5);

    console.log("=== debugFilterAndDeleteNonRoundedShapes ===");
    console.log("ID:", shape.id);
    console.log("Name:", shape.name);
    console.log("Type:", shape.type);
    console.log("Adjustments.count:", count);
    console.log("Adjustments.values:", values);

    if (isSafeNonRounded) {
        console.log("=> Shape is considered definitely NOT rounded/semi-rounded -> DELETING.");
        try {
            shape.delete();
            await context.sync();
        } catch (e) {
            console.log("Error while deleting shape:", e);
        }
    } else {
        console.log("=> Shape is kept (candidate for normalization).");
    }

    console.log("============================================");
}

// Debug: applies applyUniformRoundedCorner to all shapes that are candidates
// according to the filter heuristic.
async function debugApplyCornerRadiusOnCandidates(context, slide, radiusPts) {
    slide.shapes.load("items");
    await context.sync();

    console.log("=== debugApplyCornerRadiusOnCandidates ===");
    console.log("Shapes found:", slide.shapes.items.length);

    for (const shape of slide.shapes.items) {
        shape.load(["id", "name", "type", "adjustments"]);
    }
    await context.sync();

    for (const shape of slide.shapes.items) {
        const isGeometric = shape.type === PowerPoint.ShapeType.geometricShape;
        const adjustments = shape.adjustments;

        let count = 0;
        let firstVal = undefined;

        if (adjustments) {
            adjustments.load("count");
            await context.sync();
            count = adjustments.count;

            if (count > 0) {
                const res = adjustments.get(0);
                await context.sync();
                firstVal = res.value;
            }
        }

        const isCandidate =
            isGeometric &&
            count === 1 &&
            typeof firstVal === "number" &&
            firstVal <= 0.5;

        console.log(`Shape ${shape.id} "${shape.name}" – candidate?`, isCandidate);

        if (!isCandidate) {
            continue;
        }

        try {
            console.log(`--> applying corner radius to Shape ${shape.id}: "${shape.name}"`);
            await applyUniformRoundedCorner(context, shape, radiusPts, true, {
                shapesModified: 0,
                shapesSkipped: 0,
                shapesWithAdjustment: 0
            });
        } catch (e) {
            console.log(`--> ERROR applying to Shape ${shape.id}:`, e);
        }
    }

    console.log("=== debugApplyCornerRadiusOnCandidates done ===");
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
    shape.load(["id", "type", "groupItems", "adjustments"]);
    await context.sync();

    // Recurse into groups
    if (shape.type === PowerPoint.ShapeType.group && shape.groupItems) {
        const groupItems = shape.groupItems;
        groupItems.load("items");
        await context.sync();

        if (groupItems.items) {
            for (const subShape of groupItems.items) {
                await restoreShapeAdjustmentRecursive(context, subShape, idToOldValueMap);
            }
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


/**
 * Optional test function: inserts a single roundRectangle on the first slide
 * and sets its corner radius to 0.3.
 */
async function run() {
    try {
        await PowerPoint.run(async (context) => {
            console.log(
                "PowerPointApi 1.10 supported?",
                Office.context.requirements.isSetSupported("PowerPointApi", "1.10")
            );

            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();

            if (slides.items.length === 0) {
                console.log("No slides found.");
                showNotification("Test aborted", "No slide was found in the presentation.");
                return;
            }

            const slide = slides.items[0];

            const shape = slide.shapes.addGeometricShape(
                PowerPoint.GeometricShapeType.roundRectangle
            );
            shape.left = 100;
            shape.top = 100;
            shape.width = 200;
            shape.height = 100;

            shape.load(["width", "height", "type"]);
            await context.sync();

            console.log("Shape inserted. Width:", shape.width, "Height:", shape.height);
            console.log("Shape type:", shape.type);

            const adjustments = shape.adjustments;
            console.log("adjustments object:", adjustments);

            try {
                const adj0 = adjustments.get(0);
                await context.sync();
                console.log("Corner radius before:", adj0.value);

                adjustments.set(0, 0.3);
                await context.sync();

                console.log("Corner radius after: 0.3");
                showNotification(
                    "Test successful",
                    "A shape was inserted and its corner radius was set to 0.3."
                );
            } catch (e) {
                console.error("Error accessing Adjustment(0):", e);
                showNotification(
                    "Test: No valid adjustment",
                    "Accessing Adjustment(0) failed. Your PowerPoint client may not fully support this feature."
                );
            }
        });
    } catch (error) {
        logError("Test run()", error);
    }
}