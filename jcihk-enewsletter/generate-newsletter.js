#!/usr/bin/env node
/**
 * JCI HK Newsletter Generator
 * Outputs a fetch() snippet to paste into your browser console
 *
 * Usage:
 *   node generate-newsletter.js
 *   → Open generated JS file
 *   → Copy contents
 *   → Open Mailchimp editor in browser
 *   → Paste into DevTools Console → Enter
 *
 * When split into N emails, N fetch snippets are generated (one per Mailchimp campaign).
 */

const fs = require('fs');
const path = require('path');
const papa = require('papaparse');
const readline = require('readline');

// ============================================================
// CONFIGURATION (default values, will be overridden)
// ============================================================
const CONFIG = {
  dataCenter: 'us4',
  campaignId: '13360698',  // default; overridden by prompt at runtime
  // extract from any request, only change if it is not working
  csrfToken:  '68482a96133502c049642899504e06ecaba10dc4',

  month: '',  // will prompt user
  year:  '2026',
  csvFile: './submissions.csv',  // prepare submissions export in this file

  senderName:  'Kenneth LAW',
  senderTitle: '2026 National Digital Development Director',
  senderOrg:   'JCI Hong Kong, China',
  senderPhone: '(852) 9790 5563',
  senderEmail: 'kenneth.law@jcihk.org',
  senderAddr:  '21/F., Seaview Commercial Building, 21-24 Connaught Road West, Hong Kong',
};

// ============================================================
// CSV PARSER
// ============================================================
function parseCSV(filePath) {
  const raw     = fs.readFileSync(filePath, 'utf-8').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const csv = papa.parse(raw, { header: true });
  const rawRows = csv.data;
  const colMap = {
    'Chapter': 'chapter',
    '[e-Newsletter]  Event Date': 'date',
    '[e-Newsletter] Event Name': 'name',
    '[e-Newsletter] Event Promotion Text': 'body',
    '[e-Newsletter] Text on Call to Action Button': 'cta_text',
    '[e-Newsletter] Call to Action Button Link': 'cta_link',
    '[e-Newsletter]  Event Visual': 'visuals',
  };

  const rows = rawRows.map((row) => {
    const newObj = {};
    Object.keys(row).forEach((header) => {
      const trimHeader = header.trim();
      const targetKey = colMap[trimHeader];
      if (targetKey) newObj[targetKey] = row[header];
    });
    return newObj;
  }).filter((row) => {
    return row.name && row.body;
  });
  console.log(rows);
  return rows;
}

// ============================================================
// HELPERS
// ============================================================
let _nodeId = 200;
const nextId = () => _nodeId++;
const uid    = () => Array.from({ length: 32 }, () => '0123456789abcdef'[Math.floor(Math.random() * 16)]).join('');

// Mailchimp keys editor blocks by uniqueId; regenerate for every generated email.
function assignFreshUniqueIds(value) {
  if (value === null || typeof value !== 'object') return;
  if (Array.isArray(value)) {
    value.forEach(assignFreshUniqueIds);
    return;
  }
  if (typeof value.uniqueId === 'string') {
    value.uniqueId = uid();
  }
  Object.values(value).forEach(assignFreshUniqueIds);
}

const PARA_ATTRS = (extra = {}) => ({
  fontFamily: null, preventHardLineBreaks: false, fontSize: null,
  lineHeight: null, textAlign: null, textDirection: null,
  style: null, class: null, id: null, letterSpacing: null, paragraphMargin: null,
  ...extra
});

const emptyPara = () => ({ type: 'paragraph', attrs: PARA_ATTRS() });

// ============================================================
// TEXT → PROSEMIRROR DOC
// ============================================================
function textToDoc(text) {
  const content = text.split('\n').map(line => {
    const para = { type: 'paragraph', attrs: PARA_ATTRS() };
    if (line.trim()) para.content = [{ type: 'text', text: line || ' ' }];
    return para;
  });
  return { type: 'doc', content };
}

// ============================================================
// EVENT SECTION BUILDER
// ============================================================
function buildEventSection(event) {
  const sectionId   = nextId();
  const rowId       = nextId();
  const columnId    = nextId();
  const presenterId = nextId();
  const titleId     = nextId();
  const imageId     = nextId();
  const bodyId      = nextId();
  const buttonId    = event.cta_link ? nextId() : null;

  const colChildren = [presenterId, titleId, imageId, bodyId];
  if (buttonId) colChildren.push(buttonId);

  const nodes = {
    [sectionId]: {
      type: 'FREESTYLE_SECTION', uniqueId: uid(),
      properties: {
        emailType: 'ONE_COLUMN100',
        blockPadding: { top: 12, bottom: 12, left: 0, right: 0 },
        outerBorderStyle: 'solid', outerBorderColor: '#d9c9fe', outerBorderWidth: 7,
        blockMargin: { all: 0, top: 20, bottom: 0, left: 0, right: 0 }
      },
      id: sectionId, parent: 33, children: [rowId]
    },
    [rowId]: {
      type: 'ROW', uniqueId: uid(), properties: {},
      id: rowId, parent: sectionId, children: [columnId]
    },
    [columnId]: {
      type: 'COLUMN', uniqueId: uid(), properties: { span: 12 },
      id: columnId, parent: rowId, children: colChildren
    },
    [presenterId]: {
      type: 'TEXT', uniqueId: uid(),
      properties: {
        text: {
          type: 'doc', content: [{
            type: 'paragraph',
            attrs: PARA_ATTRS({ textAlign: 'center' }),
            content: [{
              type: 'text',
              marks: [{ type: 'textColor', attrs: { color: '#707070' } }],
              text: `Presented by ${event.chapter}`
            }]
          }]
        },
        blockPadding: { top: 6, bottom: 0, right: 24, left: 24 }
      },
      id: presenterId, parent: columnId
    },
    [titleId]: {
      type: 'TEXT', uniqueId: uid(),
      properties: {
        text: {
          type: 'doc', content: [{
            type: 'heading',
            attrs: {
              level: 1, lineHeight: null, textAlign: 'center', textDirection: null,
              fontWeight: null, fontStyle: null, fontFamily: null, fontSize: null,
              style: null, class: null, id: null, letterSpacing: null
            },
            content: [{
              type: 'text',
              marks: [
                { type: 'textColor', attrs: { color: 'rgb(255, 255, 255)' } },
                { type: 'fontSize',  attrs: { size: '20px' } }
              ],
              text: event.name || " "
            }]
          }]
        },
        blockPadding: { top: 10, bottom: 10, left: 10, right: 10 },
        backgroundColor: '#864ffe', isFullBleed: true,
        padding: { vertical: 0.5 },
        blockMargin: { top: 10, bottom: 10, left: 10, right: 10 }
      },
      id: titleId, parent: columnId
    },
    [imageId]: {
      type: 'IMAGE', uniqueId: uid(),
      properties: {
        width: 1080, percentageWidth: 0.85, size: 'scale',
        src: 'https://cdn-images.mailchimp.com/template_images/email/logo-placeholder.png',
        actualWidth: 1080, actualHeight: 1080, height: 1080,
        isLargeImage: false,
        blockPadding: { top: 12, bottom: 12, left: 16, right: 16 },
        href: '', alt: `Visual for ${event.name}`,
        fileSize: 0, name: 'placeholder.png', contentId: null
      },
      id: imageId, parent: columnId
    },
    [bodyId]: {
      type: 'TEXT', uniqueId: uid(),
      properties: {
        text: textToDoc(event.body) || ' ',
        blockPadding: { top: 12, bottom: 12, right: 24, left: 24 }
      },
      id: bodyId, parent: columnId
    },
  };

  if (buttonId) {
    nodes[buttonId] = {
      type: 'BUTTON', uniqueId: uid(),
      properties: {
        text: { type: 'button', content: [{ type: 'text', text: event.cta_text || 'Register Now' }] },
        outlookWidth: 152.56, outlookHeight: 51,
        isTargetBlank: true, href: event.cta_link,
        borderRadius: 50,
        blockPadding: { top: 6, bottom: 6, left: 24, right: 24 },
        borderSize: 1, borderStyle: 'none',
        backgroundColor: '#d9c9fe', color: '#000000'
      },
      id: buttonId, parent: columnId
    };
  }

  return { nodes, sectionId };
}

// ============================================================
// STATIC BASE NODES
// ============================================================
function buildBaseNodes(month, year, row33Children, { partIndex = 1, partTotal = 1 } = {}) {
  const newsletterTitle = partTotal > 1
    ? `Monthly e-Newsletter (${month} ${year}) - Part ${partIndex}`
    : `Monthly e-Newsletter (${month} ${year})`;

  return {
    48: {
      type: 'IMAGE', uniqueId: '9711a41113921b496c581b7894babc7f',
      properties: {
        width: '100%', percentageWidth: 1, size: 'fill',
        src: 'https://mcusercontent.com/bd43fc2f0c21b498a0d3ac256/images/d90a2916-48c2-e3d6-3743-214f6bd02d69.png',
        actualWidth: 2201, actualHeight: 800, isLargeImage: true,
        fileSize: 1116308, name: 'jcihk-2026-email-header-01.08.png', contentId: 13697104,
        blockPadding: { top: 0, bottom: 0, left: 0, right: 0 },
        backgroundColor: 'transparent', isFullBleed: false,
        padding: { vertical: 0.5 }, href: '',
        uniformCorners: true, borderRadius: 0,
        borderTopLeftRadius: 0, borderTopRightRadius: 0,
        borderBottomLeftRadius: 0, borderBottomRightRadius: 0,
        alt: '', cropData: null, height: 800
      },
      id: 48, parent: 3
    },
    3: {
      type: 'ROW', uniqueId: 'f42b1415f2ed8bf3cc9f9c6974224a72',
      properties: { padding: { horizontal: 2, top: 0, bottom: 0 } },
      id: 3, parent: 4, children: [48]
    },
    4: {
      type: 'WRAPPER', uniqueId: 'f47806060acb0915fa0ec501e0bda9c8',
      properties: {
        label: 'Header', innerBackgroundColor: 'transparent',
        outerBackgroundColor: 'transparent', outerBorderBottomStyle: 'none'
      },
      id: 4, parent: 41, children: [3]
    },
    5: {
      type: 'TEXT', uniqueId: '9c70ba2ec2de3bcbdd3a919238cd10d8',
      properties: {
        textAlign: 'center', lineHeight: 1.25, hideToolbar: false, mobile: {},
        text: {
          type: 'doc', content: [
            {
              type: 'paragraph',
              attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }),
              content: [{ type: 'text', text: 'Dear National President Senator Daryl Lin, National Immediate Past President Senator Rafael Wong, JCI Officers, Past National Presidents, National Officers, Chapter Presidents, Past Presidents, Senators, and fellow JC members,' }]
            },
            {
              type: 'paragraph',
              attrs: PARA_ATTRS({ lineHeight: '0', style: 'line-height: 0; mso-line-height-alt: 0%;', class: 'mcePastedContent' })
            },
            {
              type: 'heading',
              attrs: { level: 1, lineHeight: null, textAlign: null, textDirection: null, fontWeight: null, fontStyle: null, fontFamily: null, fontSize: null, style: null, class: 'mcePastedContent', id: null, letterSpacing: null }
            },
            {
              type: 'heading',
              attrs: { level: 3, lineHeight: null, textAlign: null, textDirection: null, fontWeight: null, fontStyle: null, fontFamily: null, fontSize: null, style: null, class: 'mcePastedContent', id: null, letterSpacing: null },
              content: [{ type: 'text', text: newsletterTitle }]
            },
            emptyPara(),
            {
              type: 'paragraph', attrs: PARA_ATTRS(),
              content: [{ type: 'text', marks: [{ type: 'fontWeightNormal' }], text: 'We will send out update every month about our latest events and initiatives, please do follow us so that you would not miss any of them.' }]
            }
          ]
        }
      },
      id: 5, parent: 33
    },
    30: {
      type: 'SOCIAL_FOLLOW', uniqueId: '4247bc9f0b7383ebb7393aa0872e9f2c',
      properties: {
        displayType: 'icon', iconType: 'dark-', alignment: 'center',
        backgroundColor: '#3dbef6', isFullBleed: true, padding: { vertical: 0.5 },
        iconStyle: 'icon', iconColor: 'light', size: 24,
        blockPadding: { top: 24, bottom: 0 }, socialType: 'follow',
        networks: [
          { key: 'website',   value: { name: 'Website',   alt: 'Website icon',   baseUrl: 'website.com'   }, url: 'https://jcihk.org',                       uniqueId: 'website-e2eee566fcad4870a5926de7b41280b8',   label: 'Website'         },
          { key: 'facebook',  value: { name: 'Facebook',  alt: 'Facebook icon',  baseUrl: 'facebook.com'  }, url: 'https://www.facebook.com/JCIHongKong',     uniqueId: 'facebook-f15053ffcb77808473a638f6fd11464c',  label: 'Facebook'        },
          { key: 'instagram', value: { name: 'Instagram', alt: 'Instagram icon', baseUrl: 'instagram.com' }, url: 'instagram.com/jcihk',                     uniqueId: 'instagram-c406f2e79e9376b70406411e5e08d311', label: 'Instagram'       },
          { key: 'linkedin',  value: { name: 'LinkedIn',  alt: 'LinkedIn icon',  baseUrl: 'linkedin.com'  }, url: 'https://hk.linkedin.com/company/jcihk',    uniqueId: 'linkedin-f9c89020b66981989000bc7ad92c4c0c',   label: 'LinkedIn'        },
          { key: 'youtube',   value: { name: 'YouTube',   alt: 'YouTube icon',   baseUrl: 'youtube.com'   }, url: 'https://www.youtube.com/@jcihongkong4400', uniqueId: 'youtube-90e60cbb927d9ce9a1db4c01db5c1b64',   label: 'YouTube Channel' }
        ]
      },
      id: 30, parent: 33
    },
    31: {
      type: 'DIVIDER', uniqueId: '6320d8fc0df6913460c702656b9c9d55',
      properties: {
        blockPadding: { top: 20, bottom: 20, right: 48, left: 48 },
        backgroundColor: '#3dbef6', isFullBleed: true,
        padding: { vertical: 0.5 }, color: '#ffffff'
      },
      id: 31, parent: 33
    },
    33: {
      type: 'ROW', uniqueId: '79cfe42cba34592d146a73a4fa4de27f',
      properties: { padding: { horizontal: 2, top: 0, bottom: 0 } },
      id: 33, parent: 34, children: row33Children
    },
    34: {
      type: 'WRAPPER', uniqueId: 'e484d4869e6a477b268d9ce0c96533b6',
      properties: {
        label: 'Body', innerBackgroundColor: '#ffffff',
        outerBackgroundColor: 'transparent', innerBackgroundImage: null,
        innerBackgroundRepeat: 'no-repeat', innerBackgroundSize: 'contain',
        innerBackgroundPosition: 'center'
      },
      id: 34, parent: 41, children: [33]
    },
    50: {
      type: 'TEXT', uniqueId: 'eb32ef329462803cc6cbd620b9f188c4',
      properties: {
        textAlign: 'center', lineHeight: 1.25, hideToolbar: false, isLinked: false,
        blockPadding: { top: 12, bottom: 12, left: 24, right: 24 },
        mobile: { blockPadding: { top: 8, bottom: 8, left: 12, right: 12 }, blockMargin: { top: 0, bottom: 0, left: 0, right: 0 } },
        text: {
          type: 'doc', content: [
            emptyPara(),
            {
              type: 'paragraph',
              attrs: PARA_ATTRS({ textAlign: 'left', textDirection: 'ltr', style: 'text-align: left; direction: ltr;', class: '' }),
              content: [{ type: 'text', text: 'We hope to see you at these exciting events! Your participation and enthusiasm are what make JCI Hong Kong, China such as vibrant and impactful community.' }]
            },
            emptyPara(),
            {
              type: 'paragraph',
              attrs: PARA_ATTRS({ textAlign: 'left', textDirection: 'ltr', style: 'text-align: left; direction: ltr;', class: '' }),
              content: [{ type: 'text', text: `Feel free to reach out to us if you have any questions and suggestions. Let us make ${year} a year of growth, community service, innovation and leadership!` }]
            },
            emptyPara(),
            {
              type: 'paragraph',
              attrs: PARA_ATTRS({ textAlign: 'left', textDirection: 'ltr', style: 'text-align: left; direction: ltr;', class: '' }),
              content: [
                { type: 'text', marks: [{ type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: 'Best Regards,' },
                { type: 'hard_break', marks: [{ type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }] },
                { type: 'hard_break', marks: [{ type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }] },
                { type: 'text', marks: [{ type: 'strong' }, { type: 'textColor', attrs: { color: 'rgb(0, 94, 126)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: CONFIG.senderName }
              ]
            },
            {
              type: 'paragraph',
              attrs: PARA_ATTRS({ textAlign: 'left', textDirection: 'ltr', style: 'text-align: left; direction: ltr;', class: '' }),
              content: [
                { type: 'text', marks: [{ type: 'textColor', attrs: { color: 'rgb(0, 161, 216)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: CONFIG.senderTitle },
                { type: 'hard_break', marks: [{ type: 'textColor', attrs: { color: 'rgb(0, 151, 215)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }] },
                { type: 'text', marks: [{ type: 'textColor', attrs: { color: 'rgb(0, 161, 216)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: CONFIG.senderOrg }
              ]
            },
            emptyPara(),
            {
              type: 'table',
              attrs: { style: null, width: null, border: '0', cellspacing: '0', cellpadding: '0', class: 'cke_show_border', id: null, align: null },
              content: [{
                type: 'table_row', attrs: { align: null, valign: null, style: null, class: null, id: null },
                content: [{
                  type: 'table_cell', attrs: { colspan: 1, rowspan: 1, style: null, valign: null, width: null, height: null, class: null, id: null },
                  content: [{
                    type: 'table', attrs: { style: null, width: null, border: '0', cellspacing: '0', cellpadding: '0', class: 'cke_show_border', id: null, align: null },
                    content: [
                      { type: 'table_row', attrs: { align: null, valign: null, style: null, class: null, id: null }, content: [{ type: 'table_cell', attrs: { colspan: 1, rowspan: 1, style: 'text-align: left;', valign: null, width: null, height: null, class: null, id: null }, content: [{ type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'strong' }, { type: 'textColor', attrs: { color: 'rgb(112, 112, 112)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: 'T: ' }, { type: 'text', marks: [{ type: 'textColor', attrs: { color: 'rgb(112, 112, 112)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: CONFIG.senderPhone }] }] }] },
                      { type: 'table_row', attrs: { align: null, valign: null, style: null, class: null, id: null }, content: [{ type: 'table_cell', attrs: { colspan: 1, rowspan: 1, style: 'text-align: left;', valign: null, width: null, height: null, class: null, id: null }, content: [{ type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'strong' }, { type: 'textColor', attrs: { color: 'rgb(112, 112, 112)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: 'E: ' }, { type: 'text', marks: [{ type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: CONFIG.senderEmail }] }] }] },
                      { type: 'table_row', attrs: { align: null, valign: null, style: null, class: null, id: null }, content: [{ type: 'table_cell', attrs: { colspan: 1, rowspan: 1, style: 'text-align: left;', valign: null, width: null, height: null, class: null, id: null }, content: [{ type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'strong' }, { type: 'textColor', attrs: { color: 'rgb(112, 112, 112)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: 'A:' }, { type: 'text', marks: [{ type: 'textColor', attrs: { color: 'rgb(112, 112, 112)' } }, { type: 'fontFamily', attrs: { fontFamily: '"DM Sans", sans-serif' } }], text: ' ' + CONFIG.senderAddr }] }] }] }
                    ]
                  }]
                }]
              }]
            }
          ]
        }
      },
      id: 50, parent: 33
    },
    51: {
      type: 'TEXT', uniqueId: '6f08a79bab4fbb412b30a2103fbf424a',
      properties: {
        hideToolbar: false, isLinked: false,
        blockPadding: { top: 12, bottom: 12, right: 24, left: 24 },
        mobile: { blockPadding: { top: 8, bottom: 8, right: 12, left: 12 }, blockMargin: { top: 0, bottom: 0, left: 0, right: 0 } },
        text: {
          type: 'doc', content: [
            emptyPara(),
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: `Copyright (C) ${year} JCI Hong Kong, China. All rights reserved.` }] },
            emptyPara(),
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: 'JCI – Junior Chamber International' }] },
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: 'Visit ' }, { type: 'text', marks: [{ type: 'link', attrs: { href: 'https://www.jci.cc', title: null, target: '_blank', pageId: null, ariaCurrent: null, style: null, linkColor: '#5b5b5b', linkUnderline: null, id: null, name: null, anchorName: null } }], text: 'www.jci.cc' }, { type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: ' to learn how young people are working to create positive change.' }] },
            { type: 'horizontal_rule' },
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: 'Please consider the environment before printing this email or its attachment(s). Thank you!' }] },
            { type: 'horizontal_rule' },
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'left', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: 'The information contained in this e-mail (including any attachment) is confidential, may be privileged and is intended solely for the intended recipient(s).  If you are not the intended recipient, please notify the sender immediately, delete this e-mail and any attachment completely from your system.  Any unauthorised use, disclosure, copying, printing, forwarding or dissemination of any part of the information in this e-mail (including any attachment) is prohibited.  There is no guarantee that this e-mail (including any attachment) is secure or error free because it may have been intercepted, corrupted, lost, delayed, incomplete or amended.  Junior Chamber International Hong Kong, China Limited does not accept liability for any damage that may be caused by this e-mail and any attachment due to whatever reasons.  Furthermore, Junior Chamber International Hong Kong, China Limited does not accept responsibility and shall not be liable for the content of any e-mail transmitted by its staff that are not for its business purposes.  Any views or opinions expressed in any e-mail are solely those of the author and do not necessarily represent those of Junior Chamber International Hong Kong, China Limited. The traffic of this e-mail may be monitored by Junior Chamber International Hong Kong, China Limited, as permitted by applicable laws and regulations.' }] },
            { type: 'horizontal_rule' },
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'center', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'link', attrs: { href: '*|ARCHIVE|*', title: null, target: null, pageId: null, ariaCurrent: null, style: null, linkColor: '#5b5b5b', linkUnderline: null, id: null, name: null, anchorName: null } }], text: 'View this email in your browser' }] },
            emptyPara(),
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'center', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: 'Want to change how you receive these emails?' }] },
            { type: 'paragraph', attrs: PARA_ATTRS({ textAlign: 'center', style: 'text-align: left;', class: '' }), content: [{ type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: 'You can ' }, { type: 'text', marks: [{ type: 'link', attrs: { href: '*|UPDATE_PROFILE|*', title: null, target: null, pageId: null, ariaCurrent: null, style: null, linkColor: '#5b5b5b', linkUnderline: null, id: null, name: null, anchorName: null } }], text: 'update your preferences' }, { type: 'text', marks: [{ type: 'textColor', attrs: { color: '#5b5b5b' } }], text: ' or ' }, { type: 'text', marks: [{ type: 'link', attrs: { href: '*|UNSUB|*', title: null, target: null, pageId: null, ariaCurrent: null, style: null, linkColor: '#5b5b5b', linkUnderline: null, id: null, name: null, anchorName: null } }], text: 'unsubscribe' }] }
          ]
        }
      },
      id: 51, parent: 39
    },
    35: { type: 'IMAGE', uniqueId: '335de3c120f2d3c99602397ec1f12b11', properties: { emailType: 'LOGO', alt: 'Logo', width: 51, isLogo: true, blockPadding: { left: 0, right: 0 }, slot: 'logo', alignSelf: 'center', src: 'https://cdn-images.mailchimp.com/template_images/email/logo-placeholder.png', actualWidth: 657, actualHeight: 580, isDefaultContent: true, height: 'auto', href: '' }, id: 35, parent: 38 },
    36: { type: 'TEXT', uniqueId: '8a7e68be7e0847ca27f3da7fe48d2436', properties: { slot: 'text', text: { type: 'doc', content: [emptyPara()] }, hideToolbar: false, mobile: {}, blockPadding: { top: 0, bottom: 0, left: 0, right: 0 } }, id: 36, parent: 38 },
    37: { type: 'FREDDIE_BADGE', uniqueId: '5cc195bb143e05a66c1b3d03eab8bfa0', properties: { slot: 'freddieBadge', isDefaultContent: true, rewards_url: 'http://eepurl.com/iMUrDg', canREMOVE_MAILCHIMP_BRANDING: true }, id: 37, parent: 38 },
    38: { type: 'EMAIL_FOOTER_SECTION', uniqueId: 'b1c06f24dc63de2ff33287d38e27beba', properties: { layout: 'centered', isLogoDisabled: true, referralBadgeAccountTheme: '20', isReferralBadgeDisabled: false, spacing: 0, innerSpacing: 0, innerPadding: { vertical: 0.5 }, slots: { logo: 35, text: 36, freddieBadge: 37 }, backgroundColor: '#ffffff', isFullBleed: true, padding: { vertical: 0.5 }, blockPadding: { top: 0, bottom: 0, left: 0, right: 0 } }, id: 38, parent: 39, children: [35, 36, 37] },
    39: { type: 'ROW', uniqueId: '1b9d0f301e170ef7bce25196e816d715', properties: { padding: { horizontal: 2, top: 0, bottom: 0 } }, id: 39, parent: 40, children: [51, 38] },
    40: { type: 'WRAPPER', uniqueId: 'b39095628a533d7a837896162361efd4', properties: { label: 'Footer', innerBackgroundColor: '#ffffff', outerBackgroundColor: 'transparent' }, id: 40, parent: 41, children: [39] },
    41: {
      type: 'ROOT', uniqueId: '239da53b4bd69a0ef1d34e2bda7d96f8',
      properties: {
        richTextKey: 'email',
        globals: {
          baseSpacing: 24, backgroundColor: '#ffffff', buttonBackgroundColor: '#000000',
          buttonBorderRadius: 0, buttonBorderSize: 2, buttonBorderColor: '#000000',
          buttonColor: '#ffffff', buttonHorizontalPadding: 28, buttonVerticalPadding: 16,
          contentBackgroundColor: '#ffffff', dividerColor: '#000000',
          headingFontFamily: 'helvetica', headingTextColor: '#000000',
          heading1FontFamily: 'dm_sans', heading2FontFamily: 'dm_sans', heading3FontFamily: 'dm_sans',
          heading4FontFamily: 'helvetica', heading1TextColor: '#000000', heading2TextColor: '#000000',
          heading3TextColor: '#000000', heading4TextColor: '#000000',
          linkTextColor: '#000000', paragraphFontFamily: 'dm_sans', buttonFontFamily: 'helvetica',
          paragraphTextColor: '#000000',
          palette: ['#ffffff', '#e2f4e6', '#797979', '#476584', '#000000'],
          textSpacing: 0, heading1FontSize: 24, heading2FontSize: 22,
          mobilePaddingRight: 0, mobilePaddingLeft: 0,
          paragraphMobileFontSize: 12, paragraphMobileLineHeight: 1.25,
          heading1MobileFontSize: 18, heading2MobileFontSize: 16,
          heading3FontSize: 20, heading3MobileFontSize: 14,
          heading4MobileFontSize: 12, heading4FontSize: 15
        },
        metadata: { createdBy: 'MAILCHIMP_TEMPLATE', validation: { isValid: true, validatedOn: Date.now() } },
        paletteOverrides: {}, themeOverrides: {}
      },
      id: 41, children: [4, 34, 40]
    }
  };
}

// ============================================================
// BUILD FULL DOCUMENT
// ============================================================
function buildDocument(events, { partIndex = 1, partTotal = 1 } = {}) {
  _nodeId = 200;
  const { eventNodes, sectionIds } = events.reduce(
    (acc, event) => {
      const { nodes, sectionId } = buildEventSection(event);
      return { eventNodes: { ...acc.eventNodes, ...nodes }, sectionIds: [...acc.sectionIds, sectionId] };
    },
    { eventNodes: {}, sectionIds: [] }
  );

  const row33Children = [5, ...sectionIds, 50, 30, 31];
  const baseNodes     = buildBaseNodes(CONFIG.month, CONFIG.year, row33Children, { partIndex, partTotal });

  const docBody = { docId: 1, document: { ...baseNodes, ...eventNodes } };
  assignFreshUniqueIds(docBody);
  return docBody;
}

// Split events into N chunks; remainder entries go to the first email.
function splitEventsIntoChunks(events, parts) {
  const n = events.length;
  const k = parts;
  const base = Math.floor(n / k);
  const remainder = n % k;
  const sizes = Array.from({ length: k }, (_, i) => (i === 0 ? base + remainder : base));

  const chunks = [];
  let offset = 0;
  for (const size of sizes) {
    chunks.push(events.slice(offset, offset + size));
    offset += size;
  }
  return { chunks, sizes };
}

// ============================================================
// GENERATE FETCH SNIPPET
// ============================================================
function generateFetchSnippet(docBody, { campaignId, partIndex = 1, partTotal = 1 } = {}) {
  const cid = campaignId ?? CONFIG.campaignId;
  const partLabel = partTotal > 1 ? ` (Part ${partIndex} of ${partTotal})` : '';
  const payload = JSON.stringify({
    ...docBody,
    html:        '',
    hash:        '',
    tabId:       Math.random().toString(36).slice(2, 7),
    forceUpdate: true
  });

  // Escape backticks and backslashes in the payload for template literal safety
  const escapedPayload = payload
    .replace(/\\/g, '\\\\')
    .replace(/`/g, '\\`')
    .replace(/\$\{/g, '\\${');

  return `// =====================================================
// JCI HK Newsletter — ${CONFIG.month} ${CONFIG.year}${partLabel}
// Paste this into the browser console while on:
// https://${CONFIG.dataCenter}.admin.mailchimp.com/email/editor?id=${cid}
// =====================================================

fetch("https://${CONFIG.dataCenter}.admin.mailchimp.com/email/editor/edit?id=${cid}", {
  method: "POST",
  mode: "cors",
  credentials: "include",
  headers: {
    "accept": "application/json",
    "content-type": "text/plain;charset=UTF-8",
    "x-csrf-token": "${CONFIG.csrfToken}",
    "x-csrf-source": "patch-csrf",
    "x-requested-with": "XMLHttpRequest"
  },
  body: \`${escapedPayload}\`
})
.then(r => r.json())
.then(d => console.log("✅ Success:", d))
.catch(e => console.error("❌ Error:", e));
`;
}

// ============================================================
// VISUAL DOWNLOAD HELPERS
// ============================================================
function parseVisualLinks(visuals) {
  if (!visuals || typeof visuals !== 'string') return [];
  return visuals
    .split(',')
    .map(link => link.trim())
    .filter(Boolean);
}

function getGoogleDriveFileId(url) {
  if (!url) return null;

  const filePathMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (filePathMatch?.[1]) return filePathMatch[1];

  const openLinkMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (openLinkMatch?.[1]) return openLinkMatch[1];

  return null;
}

function getDownloadUrlFromDriveLink(url) {
  const fileId = getGoogleDriveFileId(url);
  if (!fileId) return null;
  return `https://drive.google.com/uc?export=download&id=${fileId}`;
}

function getExtensionFromContentType(contentType) {
  if (!contentType) return '';
  const normalized = contentType.split(';')[0].trim().toLowerCase();

  const typeToExt = {
    'image/jpeg': '.jpg',
    'image/jpg': '.jpg',
    'image/png': '.png',
    'image/gif': '.gif',
    'image/webp': '.webp',
    'image/svg+xml': '.svg'
  };

  return typeToExt[normalized] || '';
}

function isImageContentType(contentType) {
  if (!contentType) return false;
  return contentType.toLowerCase().startsWith('image/');
}

function sanitizeFileName(name) {
  return String(name || '')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 120);
}

async function downloadVisuals(events, outputFile) {
  const outputBaseName = path.basename(outputFile, path.extname(outputFile));
  const visualsDir = path.join(path.dirname(outputFile), outputBaseName);

  fs.mkdirSync(visualsDir, { recursive: true });

  let downloadedCount = 0;

  for (let eventIndex = 0; eventIndex < events.length; eventIndex++) {
    const event = events[eventIndex];
    const visualLinks = parseVisualLinks(event.visuals);

    for (let visualIndex = 0; visualIndex < visualLinks.length; visualIndex++) {
      const originalLink = visualLinks[visualIndex];
      const downloadUrl = getDownloadUrlFromDriveLink(originalLink);

      if (!downloadUrl) {
        console.warn(`⚠️  Skip invalid Google Drive link (event ${eventIndex + 1}, visual ${visualIndex + 1}): ${originalLink}`);
        continue;
      }

      try {
        const response = await fetch(downloadUrl, { redirect: 'follow' });
        if (!response.ok) {
          console.warn(`⚠️  Failed download (HTTP ${response.status}) for: ${originalLink}`);
          continue;
        }

        const arrayBuffer = await response.arrayBuffer();
        const fileBuffer = Buffer.from(arrayBuffer);
        const contentType = response.headers.get('content-type');
        if (!isImageContentType(contentType)) {
          throw new Error(
            `Downloaded file is not an image (content-type: ${contentType || 'missing'}) for link: ${originalLink}`
          );
        }
        const extension = getExtensionFromContentType(contentType);

        const eventName = sanitizeFileName(event.name || `event-${eventIndex + 1}`);
        const fileName = `${String(eventIndex + 1).padStart(2, '0')}-${eventName}-visual-${visualIndex + 1}${extension}`;
        const filePath = path.join(visualsDir, fileName);

        fs.writeFileSync(filePath, fileBuffer);
        downloadedCount++;
        console.log(`🖼️  Downloaded: ${filePath}`);
      } catch (error) {
        console.warn(`⚠️  Error downloading visual for event ${eventIndex + 1}: ${error.message}`);
      }
    }
  }

  return { visualsDir, downloadedCount };
}

// ============================================================
// MAIN
// ============================================================
async function prompt(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  return new Promise(resolve => {
    rl.question(question, answer => {
      rl.close();
      resolve(answer);
    });
  });
}

async function main() {
  console.log('📧  JCI HK Newsletter Generator');
  // Prompt for month
  let monthPrompt = await prompt(`Enter the month for the newsletter (e.g., June): `);
  monthPrompt = monthPrompt.trim();
  if (!monthPrompt) {
    console.log('❌  Month not entered.');
    process.exit(1);
  }
  CONFIG.month = monthPrompt;

  // Prompt for email split count
  let splitPrompt = await prompt(`Split into how many emails? [1]: `);
  splitPrompt = splitPrompt.trim();
  let splitCount = splitPrompt ? parseInt(splitPrompt, 10) : 1;
  if (!Number.isInteger(splitCount) || splitCount < 1) {
    console.log('❌  Split count must be a positive integer.');
    process.exit(1);
  }

  // Prompt for Mailchimp campaign ID(s)
  const campaignIds = [];
  if (splitCount === 1) {
    let campaignIdPrompt = await prompt(
      `Enter Mailchimp campaign ID (from editor URL, e.g. .../editor?id=13360698) [${CONFIG.campaignId}]: `
    );
    campaignIdPrompt = campaignIdPrompt.trim();
    campaignIds.push(campaignIdPrompt || CONFIG.campaignId);
  } else {
    for (let i = 1; i <= splitCount; i++) {
      let campaignIdPrompt = await prompt(
        `Enter Mailchimp campaign ID for email ${i} of ${splitCount}: `
      );
      campaignIdPrompt = campaignIdPrompt.trim();
      if (!campaignIdPrompt) {
        console.log(`❌  Campaign ID not entered for email ${i}.`);
        process.exit(1);
      }
      campaignIds.push(campaignIdPrompt);
    }
  }

  // Prompt for download photos?
  let needPhotoInput = await prompt(`Do you want to download event photos? [Y/n]: `);
  needPhotoInput = needPhotoInput.trim().toLowerCase();
  let needDownloadVisuals = true;
  if (needPhotoInput === 'n' || needPhotoInput === 'no') {
    needDownloadVisuals = false;
  }

  console.log(`    ${CONFIG.month} ${CONFIG.year}\n`);
  if (!fs.existsSync(CONFIG.csvFile)) {
    console.error(`❌  CSV not found: ${CONFIG.csvFile}`);
    process.exit(1);
  }

  const events = parseCSV(CONFIG.csvFile);
  console.log(`✅  Loaded ${events.length} event(s):`);
  events.forEach((e, i) =>
    console.log(`    ${i + 1}. [${e.chapter}] ${e.name}${e.cta_link ? ' 🔗' : ''}`)
  );

  const { chunks, sizes } = splitEventsIntoChunks(events, splitCount);
  const outFiles = [];

  if (splitCount > 1) {
    console.log(`\n✅  Split into ${splitCount} email(s):`);
    let globalIndex = 0;
    chunks.forEach((chunk, i) => {
      console.log(`    Email ${i + 1} (campaign ${campaignIds[i]}) — ${sizes[i]} event(s):`);
      chunk.forEach((e) => {
        globalIndex++;
        console.log(`        ${globalIndex}. [${e.chapter}] ${e.name}${e.cta_link ? ' 🔗' : ''}`);
      });
      if (chunk.length === 0) {
        console.log('        (no events)');
      }
    });
  }

  const baseOutFile = `./newsletter-${CONFIG.month.toLowerCase()}-${CONFIG.year}`;

  for (let i = 0; i < splitCount; i++) {
    const chunk = chunks[i];
    const docBody = buildDocument(chunk, { partIndex: i + 1, partTotal: splitCount });
    const snippet = generateFetchSnippet(docBody, {
      campaignId: campaignIds[i],
      partIndex: i + 1,
      partTotal: splitCount,
    });
    const outFile = splitCount === 1
      ? `${baseOutFile}.js`
      : `${baseOutFile}-part${i + 1}.js`;
    fs.writeFileSync(outFile, snippet, 'utf-8');
    outFiles.push({ outFile, campaignId: campaignIds[i], partIndex: i + 1, eventCount: chunk.length });
  }

  let downloadedCount = 0;
  let visualsDir = '';
  if (needDownloadVisuals) {
    const visualsOutFile = outFiles[0].outFile;
    const result = await downloadVisuals(events, visualsOutFile);
    visualsDir = result.visualsDir;
    downloadedCount = result.downloadedCount;
    console.log(`✅  Downloaded ${downloadedCount} visual(s) → ${visualsDir}`);
  } else {
    console.log(`ℹ️  Skipped downloading visuals as per user selection.`);
  }

  console.log(`\n✅  Fetch snippet${outFiles.length > 1 ? 's' : ''} saved:`);
  outFiles.forEach(({ outFile, campaignId, partIndex, eventCount }) => {
    const partNote = splitCount > 1 ? ` (email ${partIndex}, ${eventCount} event(s))` : '';
    console.log(`    → ${outFile}${partNote} — campaign ${campaignId}`);
  });

  console.log('\n📋  Next steps:');
  console.log(`    1. Open https://${CONFIG.dataCenter}.admin.mailchimp.com — ensure you are logged in`);
  console.log('    2. Open DevTools → Console tab');
  if (outFiles.length === 1) {
    console.log(`    3. Copy & paste the contents of ${outFiles[0].outFile} into the console, then hit Enter`);
  } else {
    console.log('    3. For each file below, copy & paste into the console and hit Enter:');
    outFiles.forEach(({ outFile, campaignId, partIndex }) => {
      console.log(`        • ${outFile} (email ${partIndex}, campaign ${campaignId})`);
    });
  }
  console.log('    4. Open each email editor, review, then click Save to regenerate HTML');
  console.log('    5. Replace image placeholders manually for each event');
}

main().catch((error) => {
  console.error('❌  Failed to generate newsletter:', error);
  process.exit(1);
});
