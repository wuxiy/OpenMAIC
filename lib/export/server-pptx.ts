import pptxgen from 'pptxgenjs';
import tinycolor from 'tinycolor2';

import { type AST, toAST } from '@/lib/export/html-parser';
import { toPoints, type SvgPoints } from '@/lib/export/svg-path-parser';
import type { QuizQuestion, Scene, Stage } from '@/lib/types/stage';

type PersistedSlide = Extract<Scene['content'], { type: 'slide' }>['canvas'];
type PersistedElement = PersistedSlide['elements'][number];

const DEFAULT_FONT_SIZE = 16;
const DEFAULT_FONT_FAMILY = 'Microsoft YaHei';

export interface SkippedPptxElement {
  sceneId: string;
  slideId: string;
  elementId: string;
  type: string;
  reason: 'unsupported_on_server' | 'unsupported_server_variant' | 'malformed_geometry';
}

export class UnsupportedServerPptxElementsError extends Error {
  constructor(public readonly unsupportedElements: SkippedPptxElement[]) {
    super(
      `Server PPTX export rejected ${unsupportedElements.length} unsupported element(s): ${unsupportedElements
        .map((element) => `${element.sceneId}/${element.elementId}:${element.type}`)
        .join(', ')}`,
    );
    this.name = 'UnsupportedServerPptxElementsError';
  }
}

export interface ServerPptxResult {
  bytes: Uint8Array;
  slideCount: number;
  skippedElements: SkippedPptxElement[];
}

export interface BuildServerPptxInput {
  stage: Stage;
  scenes: Scene[];
  /** Fail closed by default. Set false only when the caller explicitly accepts degraded output. */
  strict?: boolean;
}

function formatColor(value: string) {
  if (!value) return { alpha: 0, color: '#000000' };
  const color = tinycolor(value);
  const alpha = color.getAlpha();
  return {
    alpha,
    color: alpha === 0 ? '#ffffff' : color.setAlpha(1).toHexString(),
  };
}

/** Pure HTML-to-text-props mapping shared by the Node exporter. */
function formatHtml(html: string, ratioPx2Pt: number): pptxgen.TextProps[] {
  const slices: pptxgen.TextProps[] = [];
  let bullet = false;
  let indent = 0;

  const parse = (nodes: AST[], inherited: Record<string, string> = {}) => {
    for (const node of nodes) {
      const isBlock = 'tagName' in node && ['div', 'li', 'p'].includes(node.tagName);
      if (isBlock && slices.length > 0) {
        const previous = slices[slices.length - 1];
        previous.options ??= {};
        previous.options.breakLine = true;
      }

      const style = { ...inherited };
      const styleAttribute =
        'attributes' in node
          ? node.attributes.find((attribute) => attribute.key === 'style')
          : null;
      for (const declaration of styleAttribute?.value?.split(';') ?? []) {
        const match = declaration.match(/([^:]+):\s*(.+)/);
        if (match) style[match[1].trim()] = match[2].trim();
      }

      if ('tagName' in node) {
        if (node.tagName === 'em') style['font-style'] = 'italic';
        if (node.tagName === 'strong') style['font-weight'] = 'bold';
        if (node.tagName === 'sup') style['vertical-align'] = 'super';
        if (node.tagName === 'sub') style['vertical-align'] = 'sub';
        if (node.tagName === 'a') {
          style.href = node.attributes.find((attribute) => attribute.key === 'href')?.value ?? '';
        }
        if (node.tagName === 'ul') style['list-type'] = 'ul';
        if (node.tagName === 'ol') style['list-type'] = 'ol';
        if (node.tagName === 'li') bullet = true;
        if (node.tagName === 'p') {
          const dataIndent = node.attributes.find((attribute) => attribute.key === 'data-indent');
          if (dataIndent?.value) indent = Number(dataIndent.value);
        }
      }

      if ('tagName' in node && node.tagName === 'br') {
        slices.push({ text: '', options: { breakLine: true } });
      } else if ('content' in node) {
        const options: pptxgen.TextPropsOptions = {};
        if (style['font-size']) options.fontSize = Number.parseInt(style['font-size']) / ratioPx2Pt;
        if (style.color) options.color = formatColor(style.color).color;
        if (style['background-color']) {
          options.highlight = formatColor(style['background-color']).color;
        }
        if (style['text-align']) options.align = style['text-align'] as pptxgen.HAlign;
        if (style['font-weight']) options.bold = style['font-weight'] === 'bold';
        if (style['font-style']) options.italic = style['font-style'] === 'italic';
        if (style['font-family']) options.fontFace = style['font-family'];
        if (style.href) options.hyperlink = { url: style.href };
        if (style['vertical-align'] === 'super') options.superscript = true;
        if (style['vertical-align'] === 'sub') options.subscript = true;
        if (bullet && style['list-type']) {
          options.bullet =
            style['list-type'] === 'ol'
              ? {
                  type: 'number',
                  indent: (options.fontSize ?? DEFAULT_FONT_SIZE) * 1.25,
                }
              : { indent: (options.fontSize ?? DEFAULT_FONT_SIZE) * 1.25 };
          options.paraSpaceBefore = 0.1;
          bullet = false;
        }
        if (indent) {
          options.indentLevel = indent;
          indent = 0;
        }
        slices.push({
          text: node.content
            .replace(/&nbsp;/g, ' ')
            .replace(/&gt;/g, '>')
            .replace(/&lt;/g, '<')
            .replace(/&amp;/g, '&')
            .replace(/\n/g, ''),
          options,
        });
      } else if ('children' in node) {
        parse(node.children, style);
      }
    }
  };

  parse(toAST(html));
  return slices;
}

function speakerNotes(scene: Scene): string {
  return (scene.actions ?? [])
    .filter((action) => action.type === 'speech')
    .map((action) => ('text' in action && typeof action.text === 'string' ? action.text : ''))
    .filter(Boolean)
    .join('\n');
}

function quizNotes(question: QuizQuestion): string {
  return [
    question.answer?.length ? `正确答案：${question.answer.join('、')}` : '',
    question.analysis ? `解析：${question.analysis}` : '',
  ]
    .filter(Boolean)
    .join('\n');
}

function wrapCjkText(text: string, maxCharactersPerLine: number): string {
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (normalized.length <= maxCharactersPerLine) return normalized;

  const lines: string[] = [];
  let remaining = normalized;
  while (remaining.length > maxCharactersPerLine) {
    const candidate = remaining.slice(0, maxCharactersPerLine + 1);
    const naturalBreak = Math.max(
      candidate.lastIndexOf('，'),
      candidate.lastIndexOf('。'),
      candidate.lastIndexOf('；'),
      candidate.lastIndexOf('？'),
      candidate.lastIndexOf(' '),
    );
    const breakAt =
      naturalBreak >= Math.floor(maxCharactersPerLine * 0.6)
        ? naturalBreak + 1
        : maxCharactersPerLine;
    lines.push(remaining.slice(0, breakAt).trim());
    remaining = remaining.slice(breakAt).trim();
  }
  if (remaining) lines.push(remaining);
  return lines.join('\n');
}

function addQuizSlide(
  pptx: pptxgen,
  scene: Scene,
  question: QuizQuestion,
  questionIndex: number,
  questionCount: number,
) {
  const slide = pptx.addSlide();
  slide.background = { color: 'F7F2E8' };
  slide.addText(`${scene.title} · ${questionIndex + 1}/${questionCount}`, {
    x: 0.65,
    y: 0.38,
    w: 11.9,
    h: 0.5,
    fontFace: DEFAULT_FONT_FAMILY,
    fontSize: 18,
    color: '6F655A',
    bold: true,
    margin: 0,
  });
  // At 28pt a full-width CJK glyph consumes roughly one character cell. Keep
  // quiz prompts below the measured 16:9 title-box capacity instead of relying
  // on PowerPoint's locale-dependent auto-fit, which can clip a 28-character
  // Chinese prompt in server-rendered decks.
  slide.addText(wrapCjkText(question.question, 20), {
    x: 0.65,
    y: 1.05,
    w: 12,
    h: 1.15,
    fontFace: DEFAULT_FONT_FAMILY,
    fontSize: 28,
    color: '2F302D',
    bold: true,
    margin: 0,
    fit: 'shrink',
  });

  const options = question.options ?? [];
  if (options.length > 0) {
    const rowHeight = Math.min(0.85, 3.65 / options.length);
    for (let index = 0; index < options.length; index++) {
      const option = options[index];
      const y = 2.45 + index * rowHeight;
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.72,
        y,
        w: 0.52,
        h: 0.48,
        rectRadius: 0.08,
        fill: { color: '395B64' },
        line: { color: '395B64' },
      });
      slide.addText(option.value, {
        x: 0.72,
        y: y + 0.04,
        w: 0.52,
        h: 0.36,
        fontFace: DEFAULT_FONT_FAMILY,
        fontSize: 16,
        color: 'FFFFFF',
        bold: true,
        align: 'center',
        margin: 0,
      });
      slide.addText(option.label, {
        x: 1.45,
        y: y - 0.02,
        w: 10.85,
        h: 0.58,
        fontFace: DEFAULT_FONT_FAMILY,
        fontSize: 20,
        color: '2F302D',
        margin: 0,
        fit: 'shrink',
      });
    }
  } else {
    slide.addText('请结合文本证据作答：', {
      x: 0.72,
      y: 2.5,
      w: 11.5,
      h: 0.6,
      fontFace: DEFAULT_FONT_FAMILY,
      fontSize: 20,
      color: '395B64',
      bold: true,
      margin: 0,
    });
    for (let index = 0; index < 3; index++) {
      slide.addShape(pptx.ShapeType.line, {
        x: 0.85,
        y: 3.45 + index * 0.85,
        w: 11.2,
        h: 0,
        line: { color: 'B7AEA1', width: 1.2 },
      });
    }
  }

  const notes = [speakerNotes(scene), quizNotes(question)].filter(Boolean).join('\n\n');
  if (notes) slide.addNotes(notes);
}

function setLayout(pptx: pptxgen, viewportRatio: number) {
  if (viewportRatio === 0.625) pptx.layout = 'LAYOUT_16x10';
  else if (viewportRatio === 0.75) pptx.layout = 'LAYOUT_4x3';
  else pptx.layout = 'LAYOUT_16x9';
}

type ShapePoints = NonNullable<
  NonNullable<Parameters<ReturnType<pptxgen['addSlide']>['addShape']>[1]>['points']
>;

function formatPoints(
  points: SvgPoints,
  ratioPx2Inch: number,
  scale = { x: 1, y: 1 },
): ShapePoints {
  return points.map((point) => {
    if (point.close !== undefined) return { close: true };
    if (point.type === 'M') {
      return {
        x: (Number(point.x) / ratioPx2Inch) * scale.x,
        y: (Number(point.y) / ratioPx2Inch) * scale.y,
        moveTo: true,
      };
    }
    if (point.curve?.type === 'cubic') {
      return {
        x: (Number(point.x) / ratioPx2Inch) * scale.x,
        y: (Number(point.y) / ratioPx2Inch) * scale.y,
        curve: {
          type: 'cubic',
          x1: (Number(point.curve.x1) / ratioPx2Inch) * scale.x,
          y1: (Number(point.curve.y1) / ratioPx2Inch) * scale.y,
          x2: (Number(point.curve.x2) / ratioPx2Inch) * scale.x,
          y2: (Number(point.curve.y2) / ratioPx2Inch) * scale.y,
        },
      };
    }
    if (point.curve?.type === 'quadratic') {
      return {
        x: (Number(point.x) / ratioPx2Inch) * scale.x,
        y: (Number(point.y) / ratioPx2Inch) * scale.y,
        curve: {
          type: 'quadratic',
          x1: (Number(point.curve.x1) / ratioPx2Inch) * scale.x,
          y1: (Number(point.curve.y1) / ratioPx2Inch) * scale.y,
        },
      };
    }
    return {
      x: (Number(point.x) / ratioPx2Inch) * scale.x,
      y: (Number(point.y) / ratioPx2Inch) * scale.y,
    };
  }) as ShapePoints;
}

const dashTypes = {
  solid: 'solid',
  dashed: 'dash',
  dotted: 'sysDot',
} as const;

function shapeLine(
  outline: {
    color?: string;
    width?: number;
    style?: 'solid' | 'dashed' | 'dotted';
  },
  ratioPx2Pt: number,
): pptxgen.ShapeLineProps {
  const color = formatColor(outline.color ?? '#000000');
  return {
    color: color.color,
    transparency: (1 - color.alpha) * 100,
    width: (outline.width ?? 1) / ratioPx2Pt,
    dashType: dashTypes[outline.style ?? 'solid'],
  };
}

function linePath(element: Extract<PersistedElement, { type: 'line' }>): string {
  const start = element.start.join(',');
  const end = element.end.join(',');
  if (element.broken) return `M${start} L${element.broken.join(',')} L${end}`;
  if (element.broken2) {
    return `M${start} L${element.broken2[0]},${element.start[1]} L${element.broken2[0]},${element.end[1]} L${end}`;
  }
  if (element.curve) return `M${start} Q${element.curve.join(',')} ${end}`;
  if (element.cubic) {
    return `M${start} C${element.cubic[0].join(',')} ${element.cubic[1].join(',')} ${end}`;
  }
  return `M${start} L${end}`;
}

function supportReason(element: PersistedElement): SkippedPptxElement['reason'] | null {
  if (['text', 'table', 'chart', 'code'].includes(element.type)) return null;
  if (element.type === 'shape') {
    if (element.special || element.pattern) return 'unsupported_server_variant';
    return toPoints(element.path).length ? null : 'malformed_geometry';
  }
  if (element.type === 'line') {
    return toPoints(linePath(element)).length ? null : 'malformed_geometry';
  }
  return 'unsupported_on_server';
}

function findUnsupported(slideScenes: Scene[]): SkippedPptxElement[] {
  const unsupported: SkippedPptxElement[] = [];
  for (const scene of slideScenes) {
    if (scene.content.type !== 'slide') continue;
    for (const element of scene.content.canvas.elements) {
      const reason = supportReason(element);
      if (reason) {
        unsupported.push({
          sceneId: scene.id,
          slideId: scene.content.canvas.id,
          elementId: element.id,
          type: element.type,
          reason,
        });
      }
    }
  }
  return unsupported;
}

function addSlideElements(
  pptxSlide: ReturnType<pptxgen['addSlide']>,
  slide: PersistedSlide,
  ratioPx2Inch: number,
  ratioPx2Pt: number,
  skippedElements: SkippedPptxElement[],
  sceneId: string,
) {
  for (const element of slide.elements) {
    const unsupported = skippedElements.some(
      (item) => item.sceneId === sceneId && item.elementId === element.id,
    );
    if (unsupported) continue;

    if (element.type === 'shape') {
      const rawPoints = toPoints(element.path);
      const scale = {
        x: element.width / element.viewBox[0],
        y: element.height / element.viewBox[1],
      };
      const fill = element.gradient
        ? formatColor(
            tinycolor
              .mix(element.gradient.colors[0].color, element.gradient.colors.at(-1)!.color)
              .toHexString(),
          )
        : formatColor(element.fill);
      pptxSlide.addShape('custGeom' as pptxgen.ShapeType, {
        x: element.left / ratioPx2Inch,
        y: element.top / ratioPx2Inch,
        w: element.width / ratioPx2Inch,
        h: element.height / ratioPx2Inch,
        points: formatPoints(rawPoints, ratioPx2Inch, scale),
        fill: {
          color: fill.color,
          transparency: (1 - fill.alpha * (element.opacity ?? 1)) * 100,
        },
        line: element.outline?.width ? shapeLine(element.outline, ratioPx2Pt) : undefined,
        rotate: element.rotate || undefined,
        flipH: element.flipH,
        flipV: element.flipV,
      });
      if (element.text) {
        pptxSlide.addText(formatHtml(element.text.content, ratioPx2Pt), {
          x: element.left / ratioPx2Inch,
          y: element.top / ratioPx2Inch,
          w: element.width / ratioPx2Inch,
          h: element.height / ratioPx2Inch,
          fontSize: DEFAULT_FONT_SIZE / ratioPx2Pt,
          fontFace: element.text.defaultFontName ?? DEFAULT_FONT_FAMILY,
          color: formatColor(element.text.defaultColor ?? '#000000').color,
          valign: element.text.align,
          rotate: element.rotate || undefined,
        });
      }
      continue;
    }

    if (element.type === 'line') {
      const points = formatPoints(toPoints(linePath(element)), ratioPx2Inch);
      const maxX = Math.max(element.start[0], element.end[0]);
      const maxY = Math.max(element.start[1], element.end[1]);
      const color = formatColor(element.color);
      pptxSlide.addShape('custGeom' as pptxgen.ShapeType, {
        x: element.left / ratioPx2Inch,
        y: element.top / ratioPx2Inch,
        w: maxX / ratioPx2Inch,
        h: maxY / ratioPx2Inch,
        points,
        line: {
          color: color.color,
          transparency: (1 - color.alpha) * 100,
          width: element.width / ratioPx2Pt,
          dashType: dashTypes[element.style],
          beginArrowType: element.points[0] === 'arrow' ? 'arrow' : 'none',
          endArrowType: element.points[1] === 'arrow' ? 'arrow' : 'none',
        },
      });
      continue;
    }

    if (element.type === 'chart') {
      const data = element.data.series.map((values, index) => ({
        name: element.data.legends[index] ?? `Series ${index + 1}`,
        labels: element.data.labels,
        values,
      }));
      const chartType: pptxgen.CHART_NAME = {
        bar: 'bar',
        column: 'bar',
        line: 'line',
        pie: 'pie',
        ring: 'doughnut',
        area: 'area',
        radar: 'radar',
        scatter: 'scatter',
      }[element.chartType] as pptxgen.CHART_NAME;
      const chartColors = element.themeColors.map((color) => formatColor(color).color);
      const textColor = formatColor(element.textColor ?? '#000000').color;
      pptxSlide.addChart(chartType, data, {
        x: element.left / ratioPx2Inch,
        y: element.top / ratioPx2Inch,
        w: element.width / ratioPx2Inch,
        h: element.height / ratioPx2Inch,
        chartColors,
        catAxisLabelColor: textColor,
        valAxisLabelColor: textColor,
        showLegend: data.length > 1 || ['pie', 'ring'].includes(element.chartType),
        legendPos: 'b',
        barDir: element.chartType === 'column' ? 'bar' : 'col',
        barGrouping: element.options?.stack ? 'stacked' : undefined,
        lineSmooth: element.options?.lineSmooth,
        holeSize: element.chartType === 'ring' ? 60 : undefined,
      });
      continue;
    }

    if (element.type === 'table') {
      const hidden = new Set<string>();
      for (let row = 0; row < element.data.length; row++) {
        for (let column = 0; column < element.data[row].length; column++) {
          const cell = element.data[row][column];
          for (let r = row; r < row + cell.rowspan; r++) {
            for (let c = r === row ? column + 1 : column; c < column + cell.colspan; c++) {
              hidden.add(`${r}_${c}`);
            }
          }
        }
      }
      const rows: pptxgen.TableRow[] = element.data.map((row, rowIndex) =>
        row.flatMap((cell, columnIndex) => {
          if (hidden.has(`${rowIndex}_${columnIndex}`)) return [];
          const fill = cell.style?.backcolor
            ? formatColor(cell.style.backcolor)
            : element.theme
              ? formatColor(
                  rowIndex === 0 && element.theme.rowHeader
                    ? element.theme.color
                    : tinycolor(element.theme.color)
                        .setAlpha(rowIndex % 2 === 0 ? 0.3 : 0.1)
                        .toRgbString(),
                )
              : null;
          return [
            {
              text: cell.text,
              options: {
                colspan: cell.colspan,
                rowspan: cell.rowspan,
                bold: cell.style?.bold,
                italic: cell.style?.em,
                underline: cell.style?.underline ? { style: 'sng' } : undefined,
                align: cell.style?.align,
                valign: cell.vAlign ?? 'middle',
                fontFace: cell.style?.fontname ?? DEFAULT_FONT_FAMILY,
                fontSize: Number.parseInt(cell.style?.fontsize ?? '14') / ratioPx2Pt,
                color: cell.style?.color ? formatColor(cell.style.color).color : undefined,
                fill: fill
                  ? { color: fill.color, transparency: (1 - fill.alpha) * 100 }
                  : undefined,
              },
            },
          ];
        }),
      );
      pptxSlide.addTable(rows, {
        x: element.left / ratioPx2Inch,
        y: element.top / ratioPx2Inch,
        w: element.width / ratioPx2Inch,
        h: element.height / ratioPx2Inch,
        colW: element.colWidths.map((fraction) => (element.width * fraction) / ratioPx2Inch),
        border:
          element.outline.width && element.outline.color
            ? {
                type: element.outline.style === 'solid' ? 'solid' : 'dash',
                pt: element.outline.width / ratioPx2Pt,
                color: formatColor(element.outline.color).color,
              }
            : undefined,
      });
      continue;
    }

    if (element.type === 'code') {
      const lineNumberWidth = String(element.lines.length).length;
      const code = element.lines
        .map((line, index) => {
          const prefix =
            element.showLineNumbers === false
              ? ''
              : `${String(index + 1).padStart(lineNumberWidth)}  `;
          return `${prefix}${line.content}`;
        })
        .join('\n');
      const title = element.fileName ? `${element.fileName}\n` : '';
      pptxSlide.addText(`${title}${code}`, {
        x: element.left / ratioPx2Inch,
        y: element.top / ratioPx2Inch,
        w: element.width / ratioPx2Inch,
        h: element.height / ratioPx2Inch,
        fontSize: (element.fontSize ?? 14) / ratioPx2Pt,
        fontFace: 'Courier New',
        color: '#F8F8F2',
        fill: { color: '#282A36' },
        margin: 8 / ratioPx2Pt,
        breakLine: false,
        fit: 'shrink',
      });
      continue;
    }

    if (element.type !== 'text') continue;

    const options: pptxgen.TextPropsOptions = {
      x: element.left / ratioPx2Inch,
      y: element.top / ratioPx2Inch,
      w: element.width / ratioPx2Inch,
      h: element.height / ratioPx2Inch,
      fontSize: DEFAULT_FONT_SIZE / ratioPx2Pt,
      fontFace: element.defaultFontName || DEFAULT_FONT_FAMILY,
      color: formatColor(element.defaultColor || '#000000').color,
      valign: element.vAlign ?? 'top',
      margin: 10 / ratioPx2Pt,
      paraSpaceBefore: (element.paragraphSpace ?? 5) / ratioPx2Pt,
      lineSpacingMultiple: (element.lineHeight ?? 1.5) / 1.25,
      autoFit: true,
    };
    if (element.rotate) options.rotate = element.rotate;
    if (element.wordSpace) options.charSpacing = element.wordSpace / ratioPx2Pt;
    if (element.fill) {
      const fill = formatColor(element.fill);
      options.fill = {
        color: fill.color,
        transparency: (1 - fill.alpha * (element.opacity ?? 1)) * 100,
      };
    }
    if (element.opacity !== undefined) options.transparency = (1 - element.opacity) * 100;
    if (element.vertical) options.vert = 'eaVert';

    pptxSlide.addText(formatHtml(element.content, ratioPx2Pt), options);
  }
}

/**
 * Build a PPTX directly from the persisted classroom aggregate in Node.
 *
 * This module deliberately has no browser-store or implicit network dependency.
 * The server-safe subset is exported without I/O. Unsupported/media variants fail
 * closed unless the caller explicitly opts into a degraded export with strict:false.
 */
export async function buildServerPptx(input: BuildServerPptxInput): Promise<ServerPptxResult> {
  const slideScenes = input.scenes
    .filter((scene) => scene.content.type === 'slide')
    .sort((left, right) => left.order - right.order);
  const quizScenes = input.scenes
    .filter((scene) => scene.content.type === 'quiz' && scene.content.questions.length > 0)
    .sort((left, right) => left.order - right.order);
  if (slideScenes.length === 0) throw new Error('Classroom has no exportable scenes');

  const firstSlide = slideScenes[0].content.type === 'slide' ? slideScenes[0].content.canvas : null;
  if (!firstSlide) throw new Error('Classroom has no slide scenes to export');

  const skippedElements = findUnsupported(slideScenes);
  if (input.strict !== false && skippedElements.length > 0) {
    throw new UnsupportedServerPptxElementsError(skippedElements);
  }

  const pptx = new pptxgen();
  pptx.author = 'OpenMAIC';
  pptx.company = 'OpenMAIC';
  pptx.subject = input.stage.description ?? input.stage.name;
  pptx.title = input.stage.name;
  setLayout(pptx, firstSlide.viewportRatio);

  for (const scene of [...slideScenes, ...quizScenes].sort(
    (left, right) => left.order - right.order,
  )) {
    if (scene.content.type === 'quiz') {
      const questions = scene.content.questions;
      questions.forEach((question, index) =>
        addQuizSlide(pptx, scene, question, index, questions.length),
      );
      continue;
    }
    if (scene.content.type !== 'slide') continue;
    const slide = scene.content.canvas;
    const ratioPx2Inch = 96 * (slide.viewportSize / 960);
    const ratioPx2Pt = (96 / 72) * (slide.viewportSize / 960);
    const pptxSlide = pptx.addSlide();

    if (slide.background?.type === 'solid' && slide.background.color) {
      const background = formatColor(slide.background.color);
      pptxSlide.background = {
        color: background.color,
        transparency: (1 - background.alpha) * 100,
      };
    }

    const notes = speakerNotes(scene);
    if (notes) pptxSlide.addNotes(notes);
    addSlideElements(pptxSlide, slide, ratioPx2Inch, ratioPx2Pt, skippedElements, scene.id);
  }

  const output = await pptx.write({ outputType: 'uint8array' });
  if (!(output instanceof Uint8Array)) {
    throw new Error('PPTX generator returned an unexpected output type');
  }
  const quizSlideCount = quizScenes.reduce(
    (count, scene) => count + (scene.content.type === 'quiz' ? scene.content.questions.length : 0),
    0,
  );
  return { bytes: output, slideCount: slideScenes.length + quizSlideCount, skippedElements };
}
