import './style.css'
import * as XLSX from 'xlsx'

type Entry = {
  sourceRow: number
  fields: Record<string, string>
}

type Section = 'Objectives' | 'Meeting follow up' | 'Tickets'

const sections: Section[] = [
  'Objectives',
  'Meeting follow up',
  'Tickets',
]

const labels: Record<Section, string> = {
  Objectives: 'Objetivos',
  'Meeting follow up': 'Reuniões',
  Tickets: 'Tickets',
}

let active: Section = 'Objectives'
let records: Record<Section, Entry[]> = {
  Objectives: [],
  'Meeting follow up': [],
  Tickets: [],
}

const normalize = (value: string) =>
  value.trim().toLowerCase().replace(/\s+/g, ' ')

const app = document.querySelector<HTMLDivElement>('#app')!

app.innerHTML = `
  <header>
    <div>
      <span class="eyebrow">EMBEDDED SPECIALIST</span>
      <h1>Customer workspace</h1>
      <p>Objetivos, conversas e tickets em um só lugar.</p>
    </div>
    <label class="upload">
      Importar XLSX
      <input id="file" type="file" accept=".xlsx" />
    </label>
  </header>

  <p id="message" role="status">Importe a planilha individual do cliente.</p>
  <nav aria-label="Seções"></nav>
  <main id="content"></main>
`

const content = document.querySelector<HTMLElement>('#content')!
const message = document.querySelector<HTMLElement>('#message')!
const nav = document.querySelector<HTMLElement>('nav')!
const input = document.querySelector<HTMLInputElement>('#file')!

// Dados da planilha são inseridos como texto, nunca como HTML.
function element(
  tag: string,
  text: string,
  className = '',
): HTMLElement {
  const node = document.createElement(tag)
  node.textContent = text
  node.className = className
  return node
}

function getField(entry: Entry, name: string): string {
  const match = Object.entries(entry.fields).find(
    ([key]) => normalize(key).startsWith(normalize(name)),
  )
  return match?.[1] || ''
}

function titleOf(entry: Entry): string {
  if (active === 'Objectives') return getField(entry, 'Objective')
  if (active === 'Meeting follow up') {
    return getField(entry, 'Topic discussed')
  }
  return getField(entry, 'Description')
}

function isSolved(entry: Entry): boolean {
  return ['sim', 'true'].includes(
    normalize(getField(entry, 'Solved?')),
  )
}

function objectiveRank(entry: Entry): number {
  const level = normalize(getField(entry, 'Acceptance Level'))
  if (level === 'must have') return 0
  if (level === 'highly desirable') return 1
  if (level === 'new opportunity') return 2
  return 3
}

function badge(text: string, tone = 'neutral'): HTMLElement {
  return element('span', text, `badge ${tone}`)
}

function addField(parent: HTMLElement, label: string, value: string) {
  if (!value) return

  const field = element('div', '', 'field')
  field.append(element('strong', label))

  if (/^https?:\/\/\S+$/i.test(value)) {
    const link = document.createElement('a')
    link.href = value
    link.textContent = value
    link.target = '_blank'
    link.rel = 'noopener noreferrer'
    field.append(link)
  } else {
    field.append(element('p', value))
  }

  parent.append(field)
}

function renderCard(entry: Entry): HTMLElement {
  const details = document.createElement('details')
  details.className = 'record'

  const summary = document.createElement('summary')
  const heading = element('div', '', 'record-heading')
  const badges = element('div', '', 'badges')

  if (active === 'Objectives') {
    const rank = objectiveRank(entry)
    const tones = ['critical', 'important', 'opportunity', 'neutral']
    details.classList.add('objective', `rank-${rank}`)

    badges.append(
      badge(
        getField(entry, 'Acceptance Level') || 'Sem classificação',
        tones[rank],
      ),
    )

    const status = getField(entry, 'Status')
    if (status) badges.append(badge(status))

    // Objetivos ficam abertos; o usuário pode recolher cada um.
    details.open = true
  }

  if (active === 'Meeting follow up') {
    const solved = isSolved(entry)
    const recorded = getField(entry, 'Solved?')

    details.classList.add(
      solved ? 'resolved' : recorded ? 'pending' : 'unknown',
    )

    badges.append(
      badge(
        solved ? '✓ Resolvido' : recorded ? '● Pendente' : 'Sem status',
        solved ? 'success' : recorded ? 'warning' : 'neutral',
      ),
    )
  }

  if (active === 'Tickets') {
    const type = getField(entry, 'Type')
    const isBug = normalize(type) === 'bug'

    details.classList.add(isBug ? 'bug-record' : 'feature-record')
    badges.append(
      badge(type || 'Sem tipo', isBug ? 'bug' : 'feature'),
    )

    const status = getField(entry, 'Status')
    const priority = getField(entry, 'Priority/Impact')

    if (status) {
      badges.append(
        badge(status, normalize(status) === 'done' ? 'success' : 'neutral'),
      )
    }

    if (priority) {
      const tone = /extremely|very high/i.test(priority)
        ? 'critical'
        : /high/i.test(priority)
          ? 'warning'
          : 'neutral'

      badges.append(badge(priority, tone))
    }
  }

  heading.append(badges, element('h3', titleOf(entry) || 'Sem descrição'))
  summary.append(heading, element('span', '⌄', 'chevron'))
  details.append(summary)

  const body = element('div', '', 'record-body')

  for (const [label, value] of Object.entries(entry.fields)) {
    const key = normalize(label)

    const hidden =
      active === 'Objectives'
        ? ['objective', 'target date', 'acceptance level', 'status']
        : active === 'Meeting follow up'
          ? ['meeting date', 'topic discussed', 'solved?']
          : [
              'description',
              'related goal',
              'type',
              'status',
              'priority/impact',
            ]

    if (hidden.some(prefix => key.startsWith(prefix))) continue
    addField(body, label, value)
  }

  body.append(
    element('small', `${active} · linha ${entry.sourceRow}`, 'source'),
  )

  details.append(body)
  return details
}

let selectedDate = ''
let topicFilter = 'all'
let ticketFilter = 'all'
let searchText = ''

function createSelect(
  label: string,
  options: [string, string][],
  value: string,
  onChange: (value: string) => void,
): HTMLElement {
  const wrapper = element('label', '', 'control')
  wrapper.append(element('span', label))

  const select = document.createElement('select')

  for (const [optionValue, text] of options) {
    const option = document.createElement('option')
    option.value = optionValue
    option.textContent = text
    select.append(option)
  }

  select.value = value
  select.onchange = () => onChange(select.value)
  wrapper.append(select)

  return wrapper
}

function render() {
  nav.replaceChildren()

  for (const section of sections) {
    const button = document.createElement('button')
    button.textContent = labels[section]
    button.className = section === active ? 'active' : ''
    button.setAttribute('aria-pressed', String(section === active))

    button.onclick = () => {
      active = section
      searchText = ''
      render()
    }

    nav.append(button)
  }

  content.replaceChildren()
  const entries = records[active]

  if (!entries.length) {
    content.append(
      element('p', 'Importe uma planilha para visualizar os registros.', 'empty'),
    )
    return
  }

  const toolbar = element('div', '', 'toolbar')

  const search = document.createElement('input')
  search.type = 'search'
  search.placeholder = 'Buscar nesta seção…'
  search.setAttribute('aria-label', 'Buscar nesta seção')
  search.className = 'search'
  search.value = searchText

  const list = element(
    'div',
    '',
    active === 'Objectives' ? 'objective-grid' : 'record-list',
  )

  let meetingDates: string[] = []

  if (active === 'Meeting follow up') {
    // Datas importadas no formato YYYY-MM-DD.
    meetingDates = [
      ...new Set(entries.map(entry =>
        getField(entry, 'Meeting Date') || 'Sem data',
      )),
    ].sort((a, b) => b.localeCompare(a))

    if (!meetingDates.includes(selectedDate)) {
      selectedDate = meetingDates[0]!
    }

    toolbar.append(
      createSelect(
        'Data da reunião',
        meetingDates.map(date => [date, date]),
        selectedDate,
        value => {
          selectedDate = value
          drawList()
        },
      ),
      createSelect(
        'Tópicos',
        [
          ['all', 'Todos'],
          ['pending', 'Pendentes'],
          ['resolved', 'Resolvidos'],
          ['unknown', 'Sem status'],
        ],
        topicFilter,
        value => {
          topicFilter = value
          drawList()
        },
      ),
    )
  }

  if (active === 'Tickets') {
    toolbar.append(
      createSelect(
        'Tipo de ticket',
        [
          ['all', 'Todos'],
          ['bug', 'Bugs'],
          ['feature', 'Features'],
        ],
        ticketFilter,
        value => {
          ticketFilter = value
          drawList()
        },
      ),
    )
  }

  toolbar.append(search)

  const panelHeading = element('div', '', 'panel-heading')
  const panel = element('section', '', 'data-panel')
  panel.append(panelHeading, list)
  content.append(toolbar, panel)

  function drawList() {
    list.replaceChildren()
    panelHeading.replaceChildren()

    let visible = [...entries]

    if (active === 'Objectives') {
      visible.sort((a, b) => objectiveRank(a) - objectiveRank(b))
    }

    if (active === 'Meeting follow up') {
      visible = visible.filter(entry =>
        (getField(entry, 'Meeting Date') || 'Sem data') === selectedDate,
      )

      panelHeading.append(
        element('h2', `Tópicos de ${selectedDate}`),
        element(
          'p',
          'Abra um tópico para consultar o contexto e os encaminhamentos.',
        ),
      )

      if (topicFilter === 'resolved') {
        visible = visible.filter(isSolved)
      } else if (topicFilter === 'pending') {
        visible = visible.filter(entry =>
          Boolean(getField(entry, 'Solved?')) && !isSolved(entry),
        )
      } else if (topicFilter === 'unknown') {
        visible = visible.filter(entry => !getField(entry, 'Solved?'))
      }
    }

    if (active === 'Tickets' && ticketFilter !== 'all') {
      visible = visible.filter(entry =>
        normalize(getField(entry, 'Type')) === ticketFilter,
      )
    }

    if (searchText.trim()) {
      const query = normalize(searchText)
      visible = visible.filter(entry => {
        // Busca somente em campos apresentados nesta seção.
        const values = Object.entries(entry.fields)
          .filter(([key]) =>
            !(active === 'Tickets' &&
              normalize(key).startsWith('related goal')) &&
            !(active === 'Objectives' &&
              normalize(key).startsWith('target date')),
          )
          .map(([, value]) => value)
          .join(' ')

        return normalize(values).includes(query)
      })
    }

    for (const entry of visible) {
      list.append(renderCard(entry))
    }

    if (!visible.length) {
      list.append(element('p', 'Nenhum registro para este filtro.', 'empty'))
    }
  }

  search.oninput = () => {
    searchText = search.value
    drawList()
  }

  drawList()
}

function readSection(
  workbook: XLSX.WorkBook,
  section: Section,
): Entry[] {
  const sheetName = workbook.SheetNames.find(
    name => normalize(name) === normalize(section),
  )

  if (!sheetName) {
    throw new Error(`A aba "${section}" não foi encontrada.`)
  }

  const sheet = workbook.Sheets[sheetName]!
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    blankrows: true,
    raw: false,
    dateNF: 'yyyy-mm-dd',
  })

  // Usa somente os campos do acompanhamento, sem tabelas auxiliares.
  const limits: Record<Section, number> = {
    Objectives: 6,
    'Meeting follow up': 8,
    Tickets: 9,
  }

  const headers = (rows[0] || [])
    .slice(0, limits[section])
    .map(value => String(value).trim().replace(/\s+/g, ' '))

  const requiredTitle = section === 'Meeting follow up' ? 1 : 0

  if (!headers[requiredTitle]) {
    throw new Error(`Cabeçalhos inválidos na aba "${section}".`)
  }

  const result: Entry[] = []

  for (let index = 1; index < rows.length; index++) {
    const row = rows[index]!
    // Ignora linhas vazias que contenham apenas checkbox ou formatação.
    if (!String(row[requiredTitle] ?? '').trim()) continue

    const fields: Record<string, string> = {}

    headers.forEach((header, column) => {
      if (!header || /progress\s*%/i.test(header)) return

      let value = String(row[column] ?? '').trim()

      if (normalize(header) === 'solved?') {
        if (value.toLowerCase() === 'true') value = 'Sim'
        if (value.toLowerCase() === 'false') value = 'Não'
      }

      fields[header] = value
    })

    result.push({ sourceRow: index + 1, fields })
  }

  return result
}

input.addEventListener('change', async () => {
  const file = input.files?.[0]
  if (!file) return

  input.disabled = true
  message.textContent = 'Lendo planilha…'

  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: 'array',
      cellDates: false,
    })

    // Só substitui a importação anterior se todas as abas forem lidas.
    const next = {
      Objectives: readSection(workbook, 'Objectives'),
      'Meeting follow up': readSection(workbook, 'Meeting follow up'),
      Tickets: readSection(workbook, 'Tickets'),
    }

    records = next
    message.textContent =
      `${file.name} — ` +
      `${next.Objectives.length} objetivos, ` +
      `${next['Meeting follow up'].length} tópicos e ` +
      `${next.Tickets.length} tickets.`

    render()
  } catch (error) {
    message.textContent = error instanceof Error
      ? `Não foi possível importar: ${error.message}`
      : 'Não foi possível ler o arquivo.'
  } finally {
    input.disabled = false
    input.value = ''
  }
})

render()