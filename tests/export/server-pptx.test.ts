import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';

import type { Scene, Stage } from '@/lib/types/stage';
import { buildServerPptx, UnsupportedServerPptxElementsError } from '@/lib/export/server-pptx';

const stage: Stage = {
  id: 'stage-taohuayuanji',
  name: '《桃花源记》第一课时',
  description: '八年级语文固定验收课题',
  createdAt: 1_786_563_200_000,
  updatedAt: 1_786_563_200_000,
  languageDirective: '使用简体中文',
};

const baseSlide = {
  viewportSize: 960,
  viewportRatio: 0.5625,
  theme: {
    backgroundColor: '#F7F2E8',
    themeColors: ['#395B64'],
    fontColor: '#2F302D',
    fontName: 'Microsoft YaHei',
  },
};

const scenes = [
  {
    id: 'scene-cover',
    stageId: stage.id,
    title: '封面',
    order: 0,
    type: 'slide',
    content: {
      type: 'slide',
      canvas: {
        ...baseSlide,
        id: 'slide-cover',
        background: { type: 'solid', color: '#F7F2E8' },
        elements: [
          {
            id: 'title',
            type: 'text',
            left: 96,
            top: 150,
            width: 768,
            height: 120,
            rotate: 0,
            content:
              '<p style="font-size: 48px; color: #2F302D; text-align: center"><strong>《桃花源记》</strong></p>',
            defaultFontName: 'Microsoft YaHei',
            defaultColor: '#2F302D',
          },
          {
            id: 'subtitle',
            type: 'text',
            left: 180,
            top: 292,
            width: 600,
            height: 68,
            rotate: 0,
            content:
              '<p style="font-size: 24px; text-align: center">第一课时 · 疏通文意与理解桃花源意象</p>',
            defaultFontName: 'Microsoft YaHei',
            defaultColor: '#6F655A',
          },
        ],
      },
    },
    actions: [
      {
        id: 'speech-1',
        type: 'speech',
        text: '今天我们一起走进陶渊明笔下的桃花源。',
      },
    ],
  },
  {
    id: 'scene-objectives',
    stageId: stage.id,
    title: '学习目标',
    order: 1,
    type: 'slide',
    content: {
      type: 'slide',
      canvas: {
        ...baseSlide,
        id: 'slide-objectives',
        elements: [
          {
            id: 'objectives-title',
            type: 'text',
            left: 72,
            top: 48,
            width: 816,
            height: 72,
            rotate: 0,
            content: '<p style="font-size: 34px"><strong>学习目标</strong></p>',
            defaultFontName: 'Microsoft YaHei',
            defaultColor: '#2F302D',
          },
          {
            id: 'objectives-body',
            type: 'text',
            left: 96,
            top: 148,
            width: 760,
            height: 300,
            rotate: 0,
            content:
              '<ul><li>疏通重点文言词句</li><li>梳理“发现—进入—离开—再寻”的行踪</li><li>用原文证据理解桃花源意象</li></ul>',
            defaultFontName: 'Microsoft YaHei',
            defaultColor: '#2F302D',
          },
        ],
      },
    },
  },
  {
    id: 'scene-quiz',
    stageId: stage.id,
    title: '课堂检查题',
    order: 2,
    type: 'quiz',
    content: {
      type: 'quiz',
      questions: [
        {
          id: 'q1',
          type: 'single',
          question: '下列对“便扶向路，处处志之”中“志”的解释，正确的一项是',
          options: [
            { value: 'A', label: '老人和儿童' },
            { value: 'B', label: '成年男女' },
          ],
          answer: ['A'],
          analysis: '“黄发”指老人，“垂髫”指儿童。',
          points: 10,
          hasAnswer: true,
        },
        {
          id: 'q2',
          type: 'short_answer',
          question:
            '陶渊明在《桃花源记》中描绘了怎样的社会理想？请结合文本证据回答，至少写出两点。',
          analysis: '围绕安宁、和谐、富足作答。',
          points: 10,
        },
      ],
    },
  },
] as unknown as Scene[];

describe('buildServerPptx', () => {
  it('turns a persisted Chinese, media-free classroom into Node PPTX bytes', async () => {
    const result = await buildServerPptx({ stage, scenes });

    expect(result.bytes).toBeInstanceOf(Uint8Array);
    expect(result.bytes.byteLength).toBeGreaterThan(10_000);
    expect(Array.from(result.bytes.slice(0, 2))).toEqual([0x50, 0x4b]);
    expect(result.slideCount).toBe(4);
    expect(result.skippedElements).toEqual([]);

    const zip = await JSZip.loadAsync(result.bytes);
    expect(zip.file('ppt/slides/slide1.xml')).not.toBeNull();
    expect(zip.file('ppt/slides/slide2.xml')).not.toBeNull();
    expect(zip.file('ppt/slides/slide3.xml')).not.toBeNull();
    expect(zip.file('ppt/slides/slide4.xml')).not.toBeNull();
    expect(zip.file('ppt/slides/slide5.xml')).toBeNull();

    const allSlideXml = (
      await Promise.all([
        zip.file('ppt/slides/slide1.xml')!.async('string'),
        zip.file('ppt/slides/slide2.xml')!.async('string'),
        zip.file('ppt/slides/slide3.xml')!.async('string'),
        zip.file('ppt/slides/slide4.xml')!.async('string'),
      ])
    ).join('\n');
    const allSlideText = allSlideXml.replace(/<[^>]+>/g, '');
    expect(allSlideText).toContain('桃花源记');
    expect(allSlideText).toContain('学习目标');
    expect(allSlideText).toContain('疏通重点文言词句');
    expect(allSlideText).toContain('便扶向路');
    expect(allSlideText).toContain('正确的一项是');
    expect(allSlideText).toContain('老人和儿童');
    expect(allSlideText).toContain('陶渊明在《桃花源记》中描绘了怎样的社会理想？');
    expect(allSlideText).toContain('请结合文本证据回答，至少写出两点。');
    expect(allSlideXml).toContain('F7F2E8');
    const firstQuizSlide = await zip.file('ppt/slides/slide3.xml')!.async('string');
    expect(firstQuizSlide).toContain('中“志”的解');
    expect(firstQuizSlide).toContain('释，正确的一项是');

    const notesXml = await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string');
    expect(notesXml).toContain('今天我们一起走进陶渊明笔下的桃花源');
    const quizNotes = await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string');
    expect(quizNotes).toContain('正确答案：A');
    expect(quizNotes).toContain('黄发”指老人');
  });

  it('fails explicitly when the persisted classroom contains no slide scenes', async () => {
    await expect(
      buildServerPptx({
        stage,
        scenes: [{ ...scenes[2], content: { type: 'quiz', questions: [] } } as Scene],
      }),
    ).rejects.toThrow('Classroom has no exportable scenes');
  });

  it('fails closed on media without performing implicit network access', async () => {
    const mediaScenes = structuredClone(scenes.slice(0, 1));
    const scene = mediaScenes[0];
    if (scene.content.type !== 'slide') throw new Error('invalid fixture');
    scene.content.canvas.elements.push({
      id: 'cover-image',
      type: 'image',
    } as never);
    const fetchSpy = vi.spyOn(globalThis, 'fetch');

    let error: unknown;
    try {
      await buildServerPptx({ stage, scenes: mediaScenes });
    } catch (caught) {
      error = caught;
    }

    expect(fetchSpy).not.toHaveBeenCalled();
    expect(error).toBeInstanceOf(UnsupportedServerPptxElementsError);
    expect((error as UnsupportedServerPptxElementsError).unsupportedElements).toEqual([
      {
        sceneId: 'scene-cover',
        slideId: 'slide-cover',
        elementId: 'cover-image',
        type: 'image',
        reason: 'unsupported_on_server',
      },
    ]);
    fetchSpy.mockRestore();
  });

  it('exports the pure Node shape, line, table, chart, and code subset', async () => {
    const supportedScenes = structuredClone(scenes.slice(0, 1));
    const scene = supportedScenes[0];
    if (scene.content.type !== 'slide') throw new Error('invalid fixture');
    scene.content.canvas.elements.push(
      {
        id: 'shape-1',
        type: 'shape',
        left: 40,
        top: 400,
        width: 100,
        height: 60,
        rotate: 0,
        viewBox: [100, 60],
        path: 'M0 0 L100 0 L100 60 L0 60 Z',
        fixedRatio: false,
        fill: '#395B64',
        outline: { width: 1, color: '#2F302D', style: 'solid' },
      },
      {
        id: 'line-1',
        type: 'line',
        left: 150,
        top: 430,
        width: 2,
        start: [0, 0],
        end: [120, 0],
        style: 'solid',
        color: '#2F302D',
        points: ['', 'arrow'],
      },
      {
        id: 'table-1',
        type: 'table',
        left: 300,
        top: 380,
        width: 260,
        height: 100,
        rotate: 0,
        outline: { width: 1, color: '#2F302D', style: 'solid' },
        colWidths: [0.5, 0.5],
        cellMinHeight: 30,
        data: [
          [
            { id: 'c1', colspan: 1, rowspan: 1, text: '词句' },
            { id: 'c2', colspan: 1, rowspan: 1, text: '释义' },
          ],
        ],
      },
      {
        id: 'chart-1',
        type: 'chart',
        left: 580,
        top: 360,
        width: 160,
        height: 130,
        rotate: 0,
        chartType: 'bar',
        data: {
          labels: ['初读', '再读'],
          legends: ['完成度'],
          series: [[72, 90]],
        },
        themeColors: ['#395B64'],
        textColor: '#2F302D',
      },
      {
        id: 'code-1',
        type: 'code',
        left: 750,
        top: 360,
        width: 170,
        height: 130,
        rotate: 0,
        language: 'text',
        fileName: '课文线索.txt',
        showLineNumbers: true,
        fontSize: 12,
        lines: [{ id: 'l1', content: '发现 → 进入 → 离开 → 再寻' }],
      },
    );

    const result = await buildServerPptx({ stage, scenes: supportedScenes });

    expect(result.skippedElements).toEqual([]);
    const zip = await JSZip.loadAsync(result.bytes);
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).toContain('词句');
    expect(slideXml).toContain('课文线索.txt');
    expect(slideXml).toContain('发现 → 进入 → 离开 → 再寻');
    expect(zip.file('ppt/charts/chart1.xml')).not.toBeNull();
  });
});
