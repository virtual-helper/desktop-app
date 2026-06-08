import { describe, expect, test, vi } from 'vitest'
import { createRequire } from 'node:module'
import { buildDocx } from '../electron/main/contractHandler'

vi.mock('electron', () => ({
  ipcMain: { handle: vi.fn() },
  app: { getPath: vi.fn(() => '/tmp') },
  shell: { openPath: vi.fn() },
}))

const require = createRequire(import.meta.url)
const PizZip = require('pizzip')

function makeDocx(documentXml: string, extraFiles: Record<string, string> = {}): Buffer {
  const zip = new PizZip()
  zip.file('word/document.xml', documentXml)
  for (const [name, content] of Object.entries(extraFiles)) zip.file(name, content)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function tableWithColumns(columnCount: number): string {
  const grid = Array.from({ length: columnCount }, () => '<w:gridCol w:w="500"/>').join('')
  const cells = Array.from({ length: columnCount }, (_, i) =>
    `<w:tc><w:tcPr>${i === 0 ? '<w:vMerge w:val="restart"/>' : ''}</w:tcPr><w:p><w:r><w:t>旧${i}</w:t></w:r></w:p></w:tc>`,
  ).join('')
  return `<w:tbl><w:tblGrid>${grid}</w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`
}

describe('buildDocx appendix table replacement', () => {
  test('removes placeholder highlight background from replaced text', () => {
    const out = buildDocx(makeDocx(
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>乙方：</w:t></w:r>' +
      '<w:r><w:rPr><w:highlight w:val="yellow"/><w:shd w:fill="FFFF00"/></w:rPr><w:t>【乙方】</w:t></w:r>' +
      '<w:r><w:t>（简称：乙方）</w:t></w:r></w:p>' +
      '<w:sectPr/>' +
      '</w:body></w:document>',
    ), { 乙方名称: '云南省医药有限公司' }, { 乙方: '乙方名称' }, [])

    const outXml = new PizZip(out).file('word/document.xml').asText()
    const targetRun = (outXml.match(/<w:r>[\s\S]*?云南省医药有限公司[\s\S]*?<\/w:r>/) || [''])[0]

    expect(targetRun).toContain('云南省医药有限公司')
    expect(targetRun).not.toContain('<w:highlight')
    expect(targetRun).not.toContain('<w:shd')
  })

  test('marks footer page fields for recalculation when the document opens', () => {
    const out = buildDocx(makeDocx(
      '<w:document><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
      {
        'word/settings.xml': '<w:settings><w:zoom w:percent="100"/></w:settings>',
        'word/footer1.xml':
          '<w:ftr><w:p>' +
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText>PAGE</w:instrText></w:r>' +
          '<w:r><w:t>2</w:t></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
          '<w:fldSimple w:instr="NUMPAGES"><w:r><w:t>10</w:t></w:r></w:fldSimple>' +
          '</w:p></w:ftr>',
      },
    ), {}, {}, [])

    const zip = new PizZip(out)
    const settings = zip.file('word/settings.xml').asText()
    const footer = zip.file('word/footer1.xml').asText()

    expect(settings).toContain('<w:updateFields w:val="true"/>')
    expect(footer).toContain('<w:fldChar w:fldCharType="begin" w:dirty="true"/>')
    expect(footer).toContain('<w:fldSimple w:instr="NUMPAGES" w:dirty="true">')
  })

  test('keeps page numbering continuous across generated sections', () => {
    const out = buildDocx(makeDocx(
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>正文</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="1"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>签字</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="7"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>附录</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16840" w:h="11907" w:orient="landscape"/><w:pgNumType w:start="7"/></w:sectPr></w:pPr></w:p>' +
      '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="7"/></w:sectPr>' +
      '</w:body></w:document>',
    ), {}, {}, [])

    const outXml = new PizZip(out).file('word/document.xml').asText()
    const sects = [...outXml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)]

    expect(sects[0][0]).toContain('<w:pgNumType w:start="1"/>')
    expect(sects.slice(1).some((s) => /<w:pgNumType[^>]*w:start=/.test(s[0]))).toBe(false)
  })

  test('auto-replaces the static appendix A table inside the existing landscape section', () => {
    const xml =
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>正文</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>签字页</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>附录A:季度机会品返利分销项目 目标重点客户名单</w:t></w:r></w:p>' +
      tableWithColumns(16) +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16840" w:h="11907" w:orient="landscape"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>附录B</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr>' +
      '</w:body></w:document>'

    const out = buildDocx(makeDocx(xml), {}, {}, [{
      key: '附录A',
      headers: ['客户名称', '统签备注'],
      rows: [
        { 客户名称: '云南省医药有限公司', 统签备注: '同组' },
        { 客户名称: '云南分店', 统签备注: '同组' },
      ],
      mergeColumns: ['统签备注'],
      landscape: true,
      autoLocate: true,
    }])

    const outXml = new PizZip(out).file('word/document.xml').asText()
    const sects = [...outXml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)]

    expect(outXml).toContain('附录A:季度机会品返利分销项目')
    expect(outXml).toContain('云南省医药有限公司')
    expect(outXml).not.toContain('旧0')
    expect(sects).toHaveLength(4)
    expect(sects.map((s) => (/landscape/.test(s[0]) ? '横' : '纵')).join(' ')).toBe('纵 纵 横 纵')
  })

  test('auto-location takes precedence over a plain appendix marker outside the landscape section', () => {
    const xml =
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>正文</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="1"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>签字页</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>【附录A】</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="7"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>附录A:季度机会品返利分销项目 目标重点客户名单</w:t></w:r></w:p>' +
      tableWithColumns(16) +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16840" w:h="11907" w:orient="landscape"/><w:pgNumType w:start="7"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>附录B</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="7"/></w:sectPr>' +
      '</w:body></w:document>'

    const out = buildDocx(makeDocx(xml), {}, {}, [{
      key: '附录A',
      headers: ['客户名称'],
      rows: [{ 客户名称: '云南省医药有限公司' }],
      requireHash: false,
      landscape: true,
      autoLocate: true,
    }])

    const outXml = new PizZip(out).file('word/document.xml').asText()
    const sects = [...outXml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)]

    expect(outXml).not.toContain('【附录A】')
    expect(outXml).toContain('云南省医药有限公司')
    expect(outXml).not.toContain('旧0')
    expect(sects).toHaveLength(4)
    expect(sects.map((s) => (/landscape/.test(s[0]) ? '横' : '纵')).join(' ')).toBe('纵 纵 横 纵')
    expect(sects.slice(1).some((s) => /<w:pgNumType[^>]*w:start=/.test(s[0]))).toBe(false)
  })

  test('auto-location falls back to the unique appendix-like table in a landscape section', () => {
    const xml =
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>附录A:季度机会品返利分销项目 目标重点客户名单</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr></w:pPr></w:p>' +
      tableWithColumns(16) +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16840" w:h="11907" w:orient="landscape"/></w:sectPr></w:pPr></w:p>' +
      '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr>' +
      '</w:body></w:document>'

    const out = buildDocx(makeDocx(xml), {}, {}, [{
      key: '附录A',
      headers: ['客户名称'],
      rows: [{ 客户名称: '安徽九州通医药有限公司' }],
      landscape: true,
      autoLocate: true,
    }])

    const outXml = new PizZip(out).file('word/document.xml').asText()

    expect(outXml).toContain('安徽九州通医药有限公司')
    expect(outXml).not.toContain('旧0')
  })

  test('auto-location falls back to marker insertion when template has no landscape appendix table', () => {
    const xml =
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>正文</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="1"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:br w:type="page"/><w:t>附录A</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>【附录A】</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgNumType w:start="7"/></w:sectPr>' +
      '</w:body></w:document>'

    const out = buildDocx(makeDocx(xml), {}, {}, [{
      key: '附录A',
      headers: ['客户名称'],
      rows: [{ 客户名称: '安徽九州通医药有限公司' }],
      requireHash: false,
      landscape: true,
      autoLocate: true,
    }])

    const outXml = new PizZip(out).file('word/document.xml').asText()
    const sects = [...outXml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)]
    const titleIdx = outXml.indexOf('附录A')
    const tableIdx = outXml.indexOf('安徽九州通医药有限公司')
    const nextTitleSect = sects.find((s) => s.index > titleIdx)

    expect(outXml).toContain('安徽九州通医药有限公司')
    expect(outXml).not.toContain('【附录A】')
    expect(sects.map((s) => (/landscape/.test(s[0]) ? '横' : '纵')).join(' ')).toBe('纵 纵 横 纵')
    expect(sects.slice(1).some((s) => /<w:pgNumType[^>]*w:start=/.test(s[0]))).toBe(false)
    expect(nextTitleSect?.[0]).toContain('orient="landscape"')
    expect(nextTitleSect?.[0]).toContain('w:left="992"')
    expect(nextTitleSect?.[0]).toContain('w:right="2155"')
    expect(outXml).toContain('<w:tblW w:w="5853" w:type="pct"/>')
    expect(outXml).toContain('<w:tblInd w:w="-714" w:type="dxa"/>')
    expect(outXml).toContain('<w:tblLayout w:type="fixed"/>')
    expect(outXml).toContain('<w:trHeight w:val="280"/>')
    expect(titleIdx).toBeLessThan(tableIdx)
    expect(outXml).not.toContain('<w:br w:type="page"/>')
  })

  test('removes blank paragraphs between generated appendix table and following appendix content', () => {
    const xml =
      '<w:document><w:body>' +
      '<w:p><w:r><w:t>附录A</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>【附录A】</w:t></w:r></w:p>' +
      '<w:p/>' +
      '<w:p><w:pPr><w:spacing w:line="460"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:spacing w:line="460"/></w:pPr></w:p>' +
      '<w:p><w:r><w:t>附录B：季度机会品返利分销项目折扣比例区间</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr>' +
      '</w:body></w:document>'

    const out = buildDocx(makeDocx(xml), {}, {}, [{
      key: '附录A',
      headers: ['客户名称'],
      rows: [{ 客户名称: '云南省医药有限公司' }],
      requireHash: false,
      landscape: true,
      autoLocate: true,
    }])

    const outXml = new PizZip(out).file('word/document.xml').asText()
    const tableEnd = outXml.indexOf('</w:tbl>') + '</w:tbl>'.length
    const appendixB = outXml.indexOf('附录B：')
    const between = outXml.slice(tableEnd, appendixB)

    expect(between.match(/<w:p\b/g) || []).toHaveLength(2)
    expect(between).not.toContain('<w:spacing')
  })
})
