import { bitable } from '@lark-base-open/js-sdk'
import JSZip from 'jszip'

const SIGNATURE_TAB_POS = 4828
const TEMPLATE_MAP = {
  '房屋租赁合同': '房屋租赁合同样板.docx',
  '管理服务合同': '管理服务合同模板.docx'
}

let allRecords = []
let currentRecord = null
let table = null

async function init() {
  const list = document.getElementById('recordList')
  try {
    list.innerHTML = '<div class="loading">正在连接多维表格...</div>'
    table = await bitable.base.getActiveTable()
    list.innerHTML = '<div class="loading">加载记录中...</div>'
    await loadRecords()
  } catch (e) {
    console.error('init error', e)
    list.innerHTML = '<div class="empty">初始化失败: ' + (e.message || e) + '<br><br>请确认已在多维表格中打开本插件。</div>'
  }
}

async function loadRecords() {
  const list = document.getElementById('recordList')
  list.innerHTML = '<div class="loading">加载记录中...</div>'
  try {
    const fields = await table.getFieldList()
    const fieldMap = {}
    await Promise.all(fields.map(async (f) => {
      const name = await f.getName()
      fieldMap[name] = f.id
    }))

    // 分页获取所有记录
    allRecords = []
    let pageToken = undefined
    let page = 0
    while (true) {
      page++
      const result = await table.getRecords({ pageSize: 200, pageToken })
      const batch = result.records.map(rec => {
        const obj = { recordId: rec.recordId }
        for (const [name, fid] of Object.entries(fieldMap)) {
          obj[name] = rec.fields[fid]
        }
        return obj
      })
      allRecords = allRecords.concat(batch)
      list.innerHTML = '<div class="loading">已加载 ' + allRecords.length + ' 条...</div>'
      if (!result.hasMore || !result.pageToken) break
      pageToken = result.pageToken
      if (page > 20) break // 安全上限
    }

    renderList(allRecords)
  } catch (e) {
    list.innerHTML = '<div class="empty">加载失败: ' + e.message + '</div>'
  }
}

function toText(v) {
  if (v == null) return ''
  if (Array.isArray(v)) {
    return v.map(item => {
      if (typeof item === 'object' && item !== null) return item.text || item.name || ''
      return String(item)
    }).join('、')
  }
  if (typeof v === 'object') return v.text || v.name || ''
  return String(v)
}

function fmtDate(v) {
  if (!v) return ''
  if (typeof v === 'number') {
    const d = new Date(v)
    return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日'
  }
  const s = String(v).slice(0, 10)
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/)
  if (m) return parseInt(m[1]) + '年' + parseInt(m[2]) + '月' + parseInt(m[3]) + '日'
  return s
}

function renderList(records) {
  const list = document.getElementById('recordList')
  if (!records.length) {
    list.innerHTML = '<div class="empty">暂无记录</div>'
    return
  }
  list.innerHTML = records.map(r => {
    const name = toText(r['企业名称'])
    const ctype = toText(r['合同类型'])
    const status = toText(r['合同生成状态'])
    const tagClass = ctype === '房屋租赁合同' ? 'tag-lease' : 'tag-manage'
    const statusClass = status === '已生成' ? 'tag-done' : 'tag-pending'
    return `<div class="record-item" data-id="${r.recordId}">
      <div class="record-name">${name || '(未命名)'}</div>
      <div class="record-meta">
        <span class="tag ${tagClass}">${ctype || '-'}</span>
        <span class="tag ${statusClass}">${status || '未生成'}</span>
      </div>
    </div>`
  }).join('')
  list.querySelectorAll('.record-item').forEach(el => {
    el.addEventListener('click', () => selectRecord(el.dataset.id))
  })
}

function selectRecord(recordId) {
  currentRecord = allRecords.find(r => r.recordId === recordId)
  if (!currentRecord) return
  document.querySelectorAll('.record-item').forEach(el => {
    el.classList.toggle('active', el.dataset.id === recordId)
  })
  renderDetail()
}

function renderDetail() {
  const detail = document.getElementById('detail')
  const r = currentRecord
  const rows = [
    ['企业名称', toText(r['企业名称'])],
    ['合同类型', toText(r['合同类型'])],
    ['地址', toText(r['地址'])],
    ['房间号', toText(r['房间号'])],
    ['使用面积', toText(r['使用面积'])],
    ['地址费', toText(r['地址费'])],
    ['租赁起始', fmtDate(r['租赁起始日期'])],
    ['租赁截止', fmtDate(r['租赁截止日期'])],
    ['收款抬头', toText(r['收款抬头'])],
    ['大写金额', toText(r['大写金额'])],
    ['收款信息', toText(r['收款信息'])],
    ['状态', toText(r['合同生成状态'])],
  ]
  detail.innerHTML = `<h2>${toText(r['企业名称'])}</h2>` +
    rows.map(([k, v]) => `<div class="detail-row"><div class="detail-label">${k}</div><div class="detail-value">${v || '-'}</div></div>`).join('') +
    `<div class="btn-group">
      <button class="btn btn-primary" id="genBtn">生成合同并回填</button>
      <button class="btn btn-secondary" id="downloadBtn">仅下载</button>
    </div>
    <div id="statusMsg"></div>`
  document.getElementById('genBtn').addEventListener('click', () => generateContract(true))
  document.getElementById('downloadBtn').addEventListener('click', () => generateContract(false))
  detail.style.display = 'block'
}

function buildMapping(r) {
  return {
    '企业名称': toText(r['企业名称']),
    '收款抬头': toText(r['收款抬头']),
    '地址': toText(r['地址']),
    '房间号': toText(r['房间号']),
    '使用面积': toText(r['使用面积']),
    '地址费': toText(r['地址费']),
    '租赁起始日期': fmtDate(r['租赁起始日期']),
    '租赁截止日期': fmtDate(r['租赁截止日期']),
    '大写金额': toText(r['大写金额']),
    '收款信息': toText(r['收款信息']),
  }
}

async function generateContract(shouldUpload) {
  const btn = document.getElementById('genBtn')
  const msg = document.getElementById('statusMsg')
  if (!currentRecord) return
  btn.disabled = true
  msg.innerHTML = '<div class="status-msg status-success">正在生成合同...</div>'
  try {
    const ctype = toText(currentRecord['合同类型'])
    const templateName = TEMPLATE_MAP[ctype]
    if (!templateName) throw new Error('未知合同类型: ' + ctype)

    const resp = await fetch('./templates/' + templateName)
    const templateBuf = await resp.arrayBuffer()
    const zip = await JSZip.loadAsync(templateBuf)
    const xml = await zip.file('word/document.xml').async('string')
    const mapping = buildMapping(currentRecord)

    let newXml = xml
    let count = 0
    for (const [ph, val] of Object.entries(mapping)) {
      const placeholder = '{' + ph + '}'
      if (newXml.includes(placeholder)) {
        const safeVal = String(val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        const lines = safeVal.split('\n')
        const replacement = lines.map((line, i) => i === 0 ? line : '<w:br/>' + line).join('')
        newXml = newXml.split(placeholder).join(replacement)
        count++
      }
    }

    // 落款对齐
    newXml = fixSignatureBlock(newXml, mapping)
    // 管理合同收款信息缩进
    if (ctype === '管理服务合同') newXml = indentAccountInfo(newXml)

    zip.file('word/document.xml', newXml)
    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    const fileName = ctype + '-' + toText(currentRecord['企业名称']) + '.docx'

    if (shouldUpload) {
      await uploadAttachment(blob, fileName)
      msg.innerHTML = '<div class="status-msg status-success">合同已生成并回填到附件字段！</div>'
    } else {
      downloadBlob(blob, fileName)
      msg.innerHTML = '<div class="status-msg status-success">合同已下载！</div>'
    }
    await loadRecords()
  } catch (e) {
    msg.innerHTML = '<div class="status-msg status-error">生成失败: ' + e.message + '</div>'
  } finally {
    btn.disabled = false
  }
}

function fixSignatureBlock(xml, mapping) {
  const d = mapping['租赁起始日期'] || ''
  const patterns = [
    { text: '甲方：乙方：', parts: ['甲方：', '乙方：'] },
    { text: '甲方代理人：乙方代理人：', parts: ['甲方代理人：', '乙方代理人：'] },
  ]
  if (d) patterns.push({ text: d + d, parts: [d, d] })

  for (const p of patterns) {
    const idx = xml.indexOf(p.text)
    if (idx === -1) continue
    // 找到包含该文本的段落
    const paraStart = xml.lastIndexOf('<w:p ', idx)
    const paraEnd = xml.indexOf('</w:p>', idx) + 6
    const oldPara = xml.slice(paraStart, paraEnd)
    // 构造新段落：制表位对齐
    const newPara = oldPara.replace(
      /<w:pPr>[\s\S]*?<\/w:pPr>/,
      '<w:pPr><w:tabs><w:tab w:val="left" w:pos="' + SIGNATURE_TAB_POS + '"/></w:tabs></w:pPr>'
    ).replace(
      new RegExp(p.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
      p.parts[0] + '<w:tab/>' + p.parts[1]
    )
    xml = xml.slice(0, paraStart) + newPara + xml.slice(paraEnd)
  }
  return xml
}

function indentAccountInfo(xml) {
  const idx = xml.indexOf('账户名称：')
  if (idx === -1) return xml
  const paraStart = xml.lastIndexOf('<w:p ', idx)
  const paraEnd = xml.indexOf('</w:p>', idx) + 6
  const oldPara = xml.slice(paraStart, paraEnd)
  const newPara = oldPara.replace(
    /<w:pPr>/,
    '<w:pPr><w:ind w:left="480" w:leftChars="200" w:firstLine="0" w:firstLineChars="0"/>'
  )
  return xml.slice(0, paraStart) + newPara + xml.slice(paraEnd)
}

async function uploadAttachment(blob, fileName) {
  const field = await table.getFieldByName('租赁合同附件')
  const file = new File([blob], fileName, { type: blob.type })
  await table.setRecord(currentRecord.recordId, {
    fields: { [field.id]: [file] }
  })
  // 更新状态
  try {
    const statusField = await table.getFieldByName('合同生成状态')
    await table.setRecord(currentRecord.recordId, {
      fields: { [statusField.id]: '已生成' }
    })
  } catch (e) { /* 状态字段可能不存在，忽略 */ }
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.click()
  URL.revokeObjectURL(url)
}

// 搜索
document.getElementById('search').addEventListener('input', (e) => {
  const q = e.target.value.trim().toLowerCase()
  if (!q) { renderList(allRecords); return }
  const filtered = allRecords.filter(r => toText(r['企业名称']).toLowerCase().includes(q))
  renderList(filtered)
})

init()
