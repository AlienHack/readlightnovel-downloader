import {
  Injectable,
  InternalServerErrorException,
  Logger,
} from '@nestjs/common';
import got from 'got';
import { parse } from 'node-html-parser';
import * as fs from 'fs';
import * as path from 'path';
import * as epub from 'epub-gen';
import * as docx from 'docx';
import {
  AlignmentType,
  BorderStyle,
  HeadingLevel,
  ImageRun,
  TableCell,
  TableRow,
} from 'docx';
import * as _ from 'lodash';

@Injectable()
export class ScraperService {
  private logger = new Logger('ScraperService');
  cleanTitle(title: string) {
    return title
      .replace(/:/g, '-')
      .replace(/ {2}/g, ' ')
      .replace(/\//g, '-')
      .replace(/\\/g, '-')
      .replace(/\?/g, '')
      .replace(/[^\u0E00-\u0E7Fa-zA-Z 0-9()\[\]!+\-]/g, '')
      .trim();
  }
  zero_padding(num, pad_length, pad_character = '0') {
    const pad_char = typeof pad_character !== 'undefined' ? pad_character : '0';
    const pad = new Array(1 + pad_length).join(pad_char);
    return (pad + num).slice(-pad.length);
  }
  async extractNovelDetail(novelUrl: string) {
    const { body } = await got.get(novelUrl, {
      retry: {
        limit: 3,
        methods: ['GET', 'POST'],
      },
    });

    const dom = parse(body);
    const title = dom.querySelector(
      'body > div:nth-child(4) > div > div > div.col-lg-8.content > div > div:nth-child(1) > div > div > h1',
    );
    //console.log(title.childNodes[0].rawText)
    const imageUrl = dom.querySelector(
      'body > div:nth-child(4) > div > div > div.col-lg-8.content > div > div:nth-child(2) > div > div.novel-left > div.novel-cover > a > img',
    );
    //console.log(imageUrl.attrs['src'])
    const author = dom.querySelector(
      'body > div:nth-child(4) > div > div > div.col-lg-8.content > div > div:nth-child(2) > div > div.novel-left > div.novel-details > div:nth-child(5) > div.novel-detail-body > ul > li > a',
    );
    //console.log(author.rawText)
    const shortDescriptionNodes = dom.querySelector(
      'body > div:nth-child(4) > div > div > div.col-lg-8.content > div > div:nth-child(2) > div > div.novel-right > div > div:nth-child(1) > div.novel-detail-body',
    );
    let shortDescription = '';
    for (const desc of shortDescriptionNodes.childNodes) {
      shortDescription += `<p>${desc.rawText}</p>`;
    }
    //console.log(shortDescription)
    const chapterNodes = dom.querySelector('#collapse-1 > div > div');
    const chapterTotalPage = chapterNodes.childNodes.length - 1;
    const chapters = [];
    let orderIndex = 1;
    for (let i = 0; i < chapterTotalPage; ++i) {
      const chaptersInPage = dom.querySelector(`#chapters_1-${i} > ul`);
      if (chaptersInPage == null || chaptersInPage.childNodes.length == 0)
        continue;
      const liNodes = chaptersInPage.querySelectorAll(`li`);
      for (const chapter of liNodes) {
        const chLink = chapter.querySelector(`a`).attrs['href'];
        chapters.push({
          order: orderIndex++,
          text: chapter.rawText,
          link: chLink,
        });
      }
    }
    return {
      title: title.childNodes[0].rawText,
      coverUrl: imageUrl.attrs['src'],
      author: author.rawText,
      shortDescription,
      chapters,
    };
  }
  async getChapterDetail(novelUrl: string) {
    const { body } = await got.get(novelUrl, {
      retry: {
        limit: 3,
        methods: ['GET', 'POST'],
      },
    });

    const dom = parse(body);
    const chapterTitle = dom.querySelector(
      'body > div:nth-child(5) > div > div > div.col-lg-8.content2 > div > div:nth-child(1) > div > div > h1',
    ).innerText;
    const chapterDetail = dom.querySelectorAll(
      'body > div:nth-child(5) > div > div > div.col-lg-8.content2 > div > div.chapter-content3 > div.desc > p, body > div:nth-child(5) > div > div > div.col-lg-8.content2 > div > div.chapter-content3 > div.desc > table',
    );
    const chapterBlocks = [];
    for (const ch of chapterDetail) {
      if (ch.rawTagName == 'p') {
        let chText = ch.innerText
          .replace(
            /<p class="hid">This chapter is scrapped from readlightnovel.org<\/p>/g,
            '',
          )
          .replace(
            /<p>This chapter is scrapped from readlightnovel.org<\/p>/g,
            '',
          )
          .replace(/This chapter is scrapped from readlightnovel.org/g, '')
          .replace(/&nbsp;/g, '');

        if (chText.includes('(vitag.Init')) {
          chText = chText.substring(0, chText.indexOf('(vitag.Init'));
        }

        if (chText.length != 0) {
          chapterBlocks.push({
            text: chText,
          });
        }
      } else if (ch.rawTagName == 'table') {
        const chText = ch.outerHTML
          .replace(
            /<p class="hid">This chapter is scrapped from readlightnovel.org<\/p>/g,
            '',
          )
          .replace(
            /<p>This chapter is scrapped from readlightnovel.org<\/p>/g,
            '',
          )
          .replace(/This chapter is scrapped from readlightnovel.org/g, '');

        chapterBlocks.push({
          text: chText,
        });
      }
    }
    return {
      title: chapterTitle,
      blocks: chapterBlocks,
    };
  }

  async generateEpubByPage(bookInfo, outputPath): Promise<any> {
    bookInfo.bookPathEpub = outputPath;
    await this.generateEpub(bookInfo);
  }

  async generateWordByPage(bookInfo, outputPath): Promise<any> {
    bookInfo.bookPathWord = outputPath;
    await this.generateWord(bookInfo);
  }

  async generateEpub(bookInfo): Promise<any> {
    const newEpubChapter = JSON.parse(JSON.stringify(bookInfo.chapters));
    const customCss = `
    @font-face {
      font-family: "THSarabunNew";
      font-style: normal;
      font-weight: normal;
      src : url("./fonts/THSarabunNew.ttf");
    }

    p { 
      font-family: "THSarabunNew";
    }

    h1 { 
      font-family: "THSarabunNew";
    }

    * { 
      font-family: "THSarabunNew";
    }
  `;

    const option = {
      title: bookInfo.title,
      author: bookInfo.author,
      publisher: bookInfo.author,
      cover: bookInfo.coverImage,
      content: newEpubChapter,
      lang: 'en',
      fonts: [path.join(__dirname, '../../fonts/THSarabunNew.ttf')],
      css: customCss,
      verbose: false,
      tocTitle: 'Table of Contents',
    };
    await new epub(option, bookInfo.bookPathEpub).promise;
  }
  async generateWord(bookInfo): Promise<any> {
    const sections = [];
    const imageBuffer = (
      await got.get(bookInfo.coverImage, { responseType: 'buffer' })
    ).body;

    // Construct Cover and ToC
    const coverImage = new ImageRun({
      data: imageBuffer,
      transformation: {
        width: 559,
        height: 794,
      },
      floating: {
        horizontalPosition: {
          offset: 0,
        },
        verticalPosition: {
          offset: 0,
        },
      },
    });

    const coverSection = {
      properties: {
        type: docx.SectionType.NEXT_PAGE,
        page: {
          margin: {
            top: 720,
            right: 720,
            bottom: 720,
            left: 720,
          },
          size: {
            width: 8390.55,
            height: 11905.51,
          },
        },
      },
      children: [
        new docx.Paragraph({
          children: [coverImage],
        }),
      ],
    };

    const tocSection = {
      properties: {
        page: {
          margin: {
            top: 720,
            right: 720,
            bottom: 720,
            left: 720,
          },
          size: {
            width: 8390.55,
            height: 11905.51,
          },
        },
      },
      children: [
        new docx.Paragraph({
          text: 'Table of Contents',
          heading: HeadingLevel.HEADING_1,
          border: {
            bottom: {
              color: 'auto',
              space: 1,
              value: 'single',
              size: 6,
            },
          },
        }),
        new docx.TableOfContents('Table of Contents', {
          hyperlink: true,
          headingStyleRange: '1-1',
        }),
      ],
    };

    const contentSection = [];

    // Construct Contents
    for (const chapter of bookInfo.chapters) {
      const paragraphs = [];
      paragraphs.push(
        new docx.Paragraph({
          text: chapter.title,
          heading: HeadingLevel.HEADING_1,
          border: {
            bottom: {
              color: 'auto',
              space: 1,
              value: 'single',
              size: 6,
            },
          },
        }),
      );
      for (const paragraph of chapter.blocks) {
        if (paragraph.text.includes('<table>')) {
          const dom = parse(paragraph.text);
          const tableHead = dom.querySelectorAll('table > thead > tr');
          const tableBody = dom.querySelectorAll('table > tbody > tr');
          const headers = [];
          const body = [];
          for (const head of tableHead) {
            const tableHeadData = head.querySelectorAll('td > p[dir=ltr] , *');
            for (const dt of tableHeadData) {
              headers.push(dt.innerText);
            }
          }
          for (const head of tableBody) {
            const tableBodyData = head.querySelectorAll('td > p[dir=ltr], *');
            for (const dt of tableBodyData) {
              body.push(dt.innerText);
            }
          }
          const tRow = [];
          for (const bd of body) {
            tRow.push(
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: {
                        size: 0,
                        style: BorderStyle.NONE,
                        color: '#FFFFFF',
                      },
                      left: {
                        size: 0,
                        style: BorderStyle.NONE,
                        color: '#FFFFFF',
                      },
                      right: {
                        size: 0,
                        style: BorderStyle.NONE,
                        color: '#FFFFFF',
                      },
                      bottom: {
                        size: 0,
                        style: BorderStyle.NONE,
                        color: '#FFFFFF',
                      },
                    },
                    children: [
                      new docx.Paragraph({
                        children: [
                          new docx.TextRun({
                            text: bd
                              .replace(/&nbsp;/g, ' ')
                              .replace(/&ldquo;/g, '"')
                              .replace(/&rdquo;/g, '"')
                              .replace(/&rsquo;/g, "'")
                              .replace(/&lsquo;/g, "'")
                              .replace(/&hellip;/g, '...')
                              .replace(/&ndash;/g, '-')
                              .replace(/&mdash;/g, '--')
                              .trim(),
                            size: 40,
                            font: 'TH Sarabun New',
                          }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                      }),
                    ],
                  }),
                ],
              }),
            );
          }
          if (tRow.length != 0) {
            paragraphs.push(
              new docx.Table({
                rows: tRow,
                borders: {
                  top: {
                    size: 1,
                    style: BorderStyle.SINGLE,
                    color: '#000',
                  },
                  left: {
                    size: 1,
                    style: BorderStyle.SINGLE,
                    color: '#000',
                  },
                  right: {
                    size: 1,
                    style: BorderStyle.SINGLE,
                    color: '#000',
                  },
                  bottom: {
                    size: 1,
                    style: BorderStyle.SINGLE,
                    color: '#000',
                  },
                },
              }),
            );
          } else {
            console.log(paragraph.text);
          }
          continue;
        }
        paragraphs.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text:
                  '\t' +
                  paragraph.text
                    .replace(/&nbsp;/g, ' ')
                    .replace(/&ldquo;/g, '"')
                    .replace(/&rdquo;/g, '"')
                    .replace(/&rsquo;/g, "'")
                    .replace(/&lsquo;/g, "'")
                    .replace(/&hellip;/g, '...')
                    .trim(),
                size: 40,
                font: 'TH Sarabun New',
              }),
            ],
            alignment: AlignmentType.JUSTIFIED,
          }),
        );
      }
      const content = {
        properties: {
          type: docx.SectionType.NEXT_PAGE,
          page: {
            margin: {
              top: 720,
              right: 720,
              bottom: 720,
              left: 720,
            },
            size: {
              width: 8390.55,
              height: 11905.51,
            },
          },
        },
        children: paragraphs,
      };

      contentSection.push(content);
    }

    sections.push(coverSection);
    sections.push(tocSection);

    for (const content of contentSection) {
      sections.push(content);
    }

    const doc = new docx.Document({
      creator: 'Created By ReadLightNovel Downloader',
      description: bookInfo.description,
      title: bookInfo.title,
      styles: {
        paragraphStyles: [
          {
            id: 'Heading1',
            name: 'Heading 1',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            run: {
              size: 70,
              bold: true,
              font: 'TH Sarabun New',
              color: '#50A8F2',
            },
            paragraph: {
              spacing: {
                after: 120,
              },
              alignment: AlignmentType.CENTER,
            },
          },
          {
            id: 'TOC1',
            name: 'toc 1',
            basedOn: 'Normal',
            next: 'Normal',
            quickFormat: true,
            paragraph: {},
            run: {
              font: 'TH Sarabun New',
              color: '#000000',
              size: 40,
            },
          },
        ],
      },
      sections: sections,
    });

    doc.Settings.addUpdateFields();
    const buffer = await docx.Packer.toBuffer(doc);
    fs.writeFileSync(bookInfo.bookPathWord, buffer);
  }
  async scrapingNovel(novelUrl: string) {
    const downloadDirectory = path.join(__dirname, '../../downloads/');
    await fs.promises.mkdir(downloadDirectory, { recursive: true });
    let bookInfo;
    try {
      this.logger.log('extracting novel information');
      bookInfo = await this.extractNovelDetail(novelUrl);
    } catch (err) {
      this.logger.error(err.stack);
      throw new InternalServerErrorException(
        'Cannot extract novel information',
      );
    }

    const novelDirectory = path.join(
      downloadDirectory,
      this.cleanTitle(bookInfo.title),
      '/',
    );

    await fs.promises.mkdir(novelDirectory, { recursive: true });

    const exportsDirectory = path.join(novelDirectory, 'exports/');
    await fs.promises.mkdir(exportsDirectory, { recursive: true });

    const rawDirectory = path.join(novelDirectory, 'raw/');
    await fs.promises.mkdir(rawDirectory, { recursive: true });

    const projectDirectory = path.join(novelDirectory, 'project/');
    await fs.promises.mkdir(projectDirectory, { recursive: true });

    const bookPathEpub = `${exportsDirectory}${this.cleanTitle(
      bookInfo.title,
    )}.epub`;
    const bookPathWord = `${exportsDirectory}${this.cleanTitle(
      bookInfo.title,
    )}.docx`;
    const bookProject = `${projectDirectory}${this.cleanTitle(
      bookInfo.title,
    )}.novel`;

    const chapters = [];

    const chapterWorker = JSON.parse(JSON.stringify(bookInfo.chapters));

    this.logger.log(`downloading novel chapters (${chapterWorker.length})`);
    while (chapterWorker.length) {
      await Promise.allSettled(
        chapterWorker.splice(0, 25).map(async (chapter) => {
          const chapterFile = `${novelDirectory}${this.zero_padding(
            chapter.order,
            5,
          )}.txt`;
          const rawFile = `${rawDirectory}${this.zero_padding(
            chapter.order,
            5,
          )}.txt`;

          if (fs.existsSync(rawFile)) {
            const rawCh = JSON.parse(fs.readFileSync(rawFile, 'utf8'));
            chapters.push(rawCh);
            return;
          }

          const chapterDetail = await this.getChapterDetail(chapter.link);

          let chapterRaw = '';
          for (const block of chapterDetail.blocks) {
            chapterRaw += `<p>${block.text.trim()}</p>`;
          }

          const chapterData = {
            order: chapter.order,
            title: chapterDetail.title,
            data: chapterRaw,
            blocks: chapterDetail.blocks,
            link: chapter.link,
          };

          fs.writeFileSync(chapterFile, chapterRaw);
          fs.writeFileSync(rawFile, JSON.stringify(chapterData));

          chapters.push(chapterData);
        }),
      );
    }

    const book = {
      title: bookInfo.title,
      coverImage: bookInfo.coverUrl,
      description: bookInfo.shortDescription,
      author: bookInfo.author,
      chapters: chapters,
      bookPathEpub: bookPathEpub,
      bookPathWord: bookPathWord,
    };

    book.chapters = _.sortBy(book.chapters, 'order');

    if (fs.existsSync(bookProject)) {
      this.logger.log(`merging existing project file`);
      const existedProject = JSON.parse(fs.readFileSync(bookProject, 'utf8'));
      book.chapters = _.unionBy(book.chapters, existedProject.chapters, 'link');
      book.chapters = _.sortBy(book.chapters, 'order');
    }

    this.logger.log(`saving project file`);
    fs.writeFileSync(bookProject, JSON.stringify(book));

    //GENERATE BOOKS
    // this.logger.log(`generating epub`);
    // await this.generateEpub(book);
    // this.logger.log(`generating docx`);
    // await this.generateWord(book);
    //END GENERATE BOOKS

    const chapterContent = book.chapters;
    const totalChapter = book.chapters.length;
    let chapterFrom = 1;
    let chapterTo = 100;
    if (chapterTo > totalChapter) {
      chapterTo = totalChapter;
    }

    while (chapterTo <= totalChapter) {
      const fileNameEpub = path.join(
        exportsDirectory,
        this.cleanTitle(book.title) + ` ${chapterFrom}-${chapterTo}.epub`,
      );

      const fileNameWord = path.join(
        exportsDirectory,
        this.cleanTitle(book.title) + ` ${chapterFrom}-${chapterTo}.docx`,
      );

      book.chapters = chapterContent.filter(
        (c) => c.order >= chapterFrom && c.order <= chapterTo,
      );

      if (!fs.existsSync(fileNameEpub)) {
        try {
          await this.generateEpubByPage(book, fileNameEpub);
        } catch (err) {
          this.logger.error(err.stack);
        }
      }

      if (!fs.existsSync(fileNameWord)) {
        try {
          await this.generateWordByPage(book, fileNameWord);
        } catch (err) {
          this.logger.error(err.stack);
        }
      }

      chapterFrom = chapterTo + 1;
      chapterTo = chapterTo + 100;

      if (chapterFrom > totalChapter) {
        break;
      }

      if (chapterTo > totalChapter) {
        chapterTo = totalChapter;
      }
    }

    return {
      success: true,
      detail: `The epub/docx has been downloaded and generated`,
    };
  }
}
