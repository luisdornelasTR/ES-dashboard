import './style.css'
import * as XLSX from 'xlsx'
import trLogo from './assets/tr-logo.png'

type Section = 'Objectives' | 'Meeting follow up' | 'Tickets'

type Entry = {
  sourceRow: number
  fields: Record<string, string>
}

const sections: Section[] = [
  'Tickets',
  'Meeting follow up',
  'Objectives',
]

const labels: Record<Section, string> = {
  Objectives: 'Objectives',
  'Meeting follow up': 'Meetings',
  Tickets: 'Tickets',
}

const normalize = (value: string) =>
  value.trim().toLowerCase().replace(/\s+/g, ' ')

let active: Section = 'Tickets'
let selectedDates: Set<string> | null = null
let selectedTopicStates: Set<string> | null = null
let selectedTicketTypes: Set<string> | null = null
let selectedTicketStatuses: Set<string> | null = null
let searchText = ''
const expandedMeetingDates = new Set<string>()

let records: Record<Section, Entry[]> = {
  Objectives: [],
  'Meeting follow up': [],
  Tickets: [],
}

const app = document.querySelector<HTMLDivElement>('#app')!

document.documentElement.lang = 'en'
document.title = 'ES Dashboard'

app.innerHTML = `
  <header class="workspace-header">
    <div class="brand">
      <img class="brand-logo" src="${trLogo}" alt="testRigor" />
      <div class="brand-copy">
        <span class="eyebrow">EMBEDDED SPECIALIST PROGRAM</span>
        <span class="brand-description">CSM Dashboard</span>
      </div>
    </div>

    <label class="upload">
      <span class="upload-icon" aria-hidden="true">↑</span>
      Import spreadsheet
      <input id="file" type="file" accept=".xlsx" />
    </label>
  </header>

  <section class="workspace-intro" aria-labelledby="workspace-title">
    <span class="section-label">CUSTOMER FOLLOW-UP</span>
    <h1 id="workspace-title">Engagement overview</h1>
    <p>Objectives, meeting discussions and technical follow-up.</p>
    <div class="title-rule" aria-hidden="true"></div>
  </section>

  <p id="message" role="status">
    Import an individual customer tracking spreadsheet.
  </p>

  <nav aria-label="Workspace sections"></nav>
  <main id="content"></main>
`

const content = document.querySelector<HTMLElement>('#content')!
const message = document.querySelector<HTMLElement>('#message')!
const nav = document.querySelector<HTMLElement>('nav')!
const input = document.querySelector<HTMLInputElement>('#file')!

function element(
  tag: string,
  text = '',
  className = '',
): HTMLElement {
  const node = document.createElement(tag)
  node.textContent = text
  node.className = className
  return node
}

function getField(entry: Entry, name: string): string {
  return Object.entries(entry.fields).find(
    ([key]) => normalize(key).startsWith(normalize(name)),
  )?.[1] || ''
}

function titleOf(entry: Entry): string {
  if (active === 'Objectives') return getField(entry, 'Objective')
  if (active === 'Meeting follow up') {
    return getField(entry, 'Topic discussed')
  }
  return getField(entry, 'Description')
}

function topicState(entry: Entry): string {
  const value = normalize(getField(entry, 'Solved?'))
  if (['true', 'yes', 'sim'].includes(value)) return 'resolved'
  if (['false', 'no', 'não'].includes(value)) return 'pending'
  return 'unknown'
}

function objectiveRank(entry: Entry): number {
  const level = normalize(getField(entry, 'Acceptance Level'))
  if (level === 'must have') return 0
  if (level === 'highly desirable') return 1
  if (level === 'new opportunity') return 2
  return 3
}

function isFinished(entry: Entry): boolean {
  return [
    'done',
    'closed',
    'resolved',
    'completed',
    'rejected',
    'cancelled',
    'canceled',
  ].includes(normalize(getField(entry, 'Status')))
}

function statusTone(value: string): string {
  const status = normalize(value)

  if (['done', 'completed', 'resolved', 'closed'].includes(status)) {
    return 'success'
  }

  if (['rejected', 'cancelled', 'canceled', 'blocked'].includes(status)) {
    return 'critical'
  }

  if (['in progress', 'in review', 'prioritized', 'delayed'].includes(status)) {
    return 'warning'
  }

  if (status === 'on track') return 'success'
  return 'neutral'
}

function badge(text: string, tone = 'neutral'): HTMLElement {
  return element('span', text, `badge ${tone}`)
}

function metric(
  label: string,
  count: number,
  hint: string,
  onClick: () => void,
  selected: boolean,
  color = '#b1d1f8',
): HTMLButtonElement {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = 'metric-card'
  button.style.setProperty('--metric-color', color)
  button.setAttribute('aria-pressed', String(selected))
  button.append(
    element('span', label, 'metric-label'),
    element('strong', String(count), 'metric-value'),
    element('span', hint, 'metric-hint'),
  )
  button.onclick = onClick
  return button
}

function topicCounts(entries: Entry[]) {
  return {
    total: entries.length,
    resolved: entries.filter(entry => topicState(entry) === 'resolved').length,
    pending: entries.filter(entry => topicState(entry) === 'pending').length,
    unknown: entries.filter(entry => topicState(entry) === 'unknown').length,
  }
}

function distribution(
  parts: { label: string; count: number; color: string; onClick?: () => void }[],
  label: string,
): HTMLElement {
  const bar = element('div', '', 'distribution-bar')
  bar.setAttribute('role', 'group')
  bar.setAttribute('aria-label', label)
  const total = parts.reduce((sum, part) => sum + part.count, 0)

  for (const part of parts.filter(part => part.count > 0)) {
    const text = `${part.label}: ${part.count} (${(part.count / total * 100).toFixed(1)}%)`
    const segment = part.onClick
      ? document.createElement('button')
      : element('span')
    segment.className = 'distribution-segment'
    segment.style.flexGrow = String(part.count)
    segment.style.background = part.color
    segment.title = text
    segment.setAttribute('aria-label', text)
    if (segment instanceof HTMLButtonElement) {
      segment.type = 'button'
      segment.onclick = part.onClick!
    }
    bar.append(segment)
  }
  return bar
}

function meetingDate(entry: Entry): string {
  return getField(entry, 'Meeting Date')
}

function formatDate(value: string): string {
  if (!value) return 'No date'

  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value)
  if (!match) return value

  const date = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
  )

  return date.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  })
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

function createMultiSelect(
  label: string,
  options: [string, string][],
  selected: Set<string>,
  onChange: (values: Set<string>) => void,
): HTMLElement {
  const wrapper = element('div', '', 'control multi-control')
  wrapper.append(element('span', label))

  const dropdown = document.createElement('details')
  dropdown.className = 'multi-select'

  const summary = document.createElement('summary')
  const summaryText = element('span')
  const arrow = element('span', '⌄', 'multi-arrow')
  arrow.setAttribute('aria-hidden', 'true')
  summary.append(summaryText, arrow)

  const panel = element('div', '', 'multi-panel')
  const actions = element('div', '', 'multi-actions')
  const selectAll = document.createElement('button')
  const clear = document.createElement('button')

  selectAll.type = 'button'
  selectAll.textContent = 'Select all'
  clear.type = 'button'
  clear.textContent = 'Clear'
  actions.append(selectAll, clear)

  const filter = document.createElement('input')
  filter.type = 'search'
  filter.className = 'multi-search'
  filter.placeholder = 'Find an option…'
  filter.setAttribute('aria-label', `Search ${label.toLowerCase()}`)

  const optionList = element('div', '', 'multi-options')
  const empty = element('p', 'No matching options.', 'multi-empty')
  empty.hidden = true

  const checkboxes: HTMLInputElement[] = []
  const optionRows: { row: HTMLElement; text: string }[] = []

  function updateSummary() {
    const chosen = options.filter(([value]) => selected.has(value))

    summaryText.textContent =
      chosen.length === 0
        ? 'None selected'
        : chosen.length === options.length
          ? 'All selected'
          : chosen.length === 1
            ? chosen[0]![1]
            : `${chosen.length} selected`
  }

  function notifyChange() {
    updateSummary()
    onChange(new Set(selected))
  }

  for (const [value, text] of options) {
    const row = document.createElement('label')
    row.className = 'multi-option'

    const checkbox = document.createElement('input')
    checkbox.type = 'checkbox'
    checkbox.value = value
    checkbox.checked = selected.has(value)

    checkbox.onchange = () => {
      if (checkbox.checked) selected.add(value)
      else selected.delete(value)
      notifyChange()
    }

    row.append(checkbox, element('span', text))
    optionList.append(row)
    checkboxes.push(checkbox)
    optionRows.push({ row, text })
  }

  selectAll.onclick = () => {
    for (const [value] of options) selected.add(value)
    checkboxes.forEach(checkbox => {
      checkbox.checked = true
    })
    notifyChange()
  }

  clear.onclick = () => {
    selected.clear()
    checkboxes.forEach(checkbox => {
      checkbox.checked = false
    })
    notifyChange()
  }

  filter.oninput = () => {
    let matches = 0
    const query = normalize(filter.value)

    for (const { row, text } of optionRows) {
      row.hidden = !normalize(text).includes(query)
      if (!row.hidden) matches++
    }

    empty.hidden = matches > 0
  }

  dropdown.addEventListener('keydown', event => {
    if (event.key === 'Escape') {
      dropdown.open = false
      summary.focus()
    }
  })

  dropdown.addEventListener('toggle', () => {
    if (!dropdown.open) return

    document.querySelectorAll<HTMLDetailsElement>('.multi-select').forEach(
      other => {
        if (other !== dropdown) other.open = false
      },
    )
  })

  panel.append(actions, filter, optionList, empty)
  dropdown.append(summary, panel)
  wrapper.append(dropdown)
  updateSummary()

  return wrapper
}

// Close filter menus when clicking outside them.
document.addEventListener('click', event => {
  if (!(event.target instanceof Node)) return
  const target = event.target

  document.querySelectorAll<HTMLDetailsElement>('.multi-select[open]').forEach(
    dropdown => {
      if (!dropdown.contains(target)) dropdown.open = false
    },
  )
})

function renderCard(entry: Entry): HTMLDetailsElement {
  const details = document.createElement('details')
  details.className = 'record'
  details.dataset.row = String(entry.sourceRow)

  const summary = document.createElement('summary')
  const heading = element('div', '', 'record-heading')
  const badges = element('div', '', 'badges')
  const body = element('div', '', 'record-body')

  if (active === 'Objectives') {
    const rank = objectiveRank(entry)
    const tones = ['critical', 'important', 'opportunity', 'neutral']

    details.classList.add('objective', `rank-${rank}`)
    details.open = true

    badges.append(
      badge(
        getField(entry, 'Acceptance Level') || 'Unclassified',
        tones[rank],
      ),
    )

    const status = getField(entry, 'Status')
    badges.append(
      badge(`Status: ${status || 'Not specified'}`, statusTone(status)),
    )
  }

  if (active === 'Meeting follow up') {
    const state = topicState(entry)
    details.classList.add(state)

    const text = state === 'resolved'
      ? '✓ Resolved'
      : state === 'pending'
        ? '● Pending'
        : 'Status not recorded'

    badges.append(
      badge(
        text,
        state === 'resolved'
          ? 'success'
          : state === 'pending'
            ? 'warning'
            : 'neutral',
      ),
    )
  }

  if (active === 'Tickets') {
    const type = getField(entry, 'Type')
    const isBug = normalize(type) === 'bug'
    const isFeature = normalize(type) === 'feature'

    if (isBug) details.classList.add('bug-record')
    if (isFeature) details.classList.add('feature-record')

    badges.append(
      badge(
        type || 'No type',
        isBug ? 'bug' : isFeature ? 'feature' : 'neutral',
      ),
    )

    const status = getField(entry, 'Status')
    badges.append(
      badge(status || 'No status', statusTone(status)),
    )

    const priority = getField(entry, 'Priority/Impact')
    if (priority) {
      badges.append(
        badge(
          priority,
          /extremely|very high/i.test(priority)
            ? 'critical'
            : /high/i.test(priority)
              ? 'warning'
              : 'neutral',
        ),
      )
    }

    const jira = getField(entry, 'Jira Link')
    const ticketId = jira.match(/\/browse\/([A-Z]+-\d+)/)?.[1]
    if (ticketId) badges.append(badge(ticketId))
  }

  heading.append(
    badges,
    element('h3', titleOf(entry) || 'Untitled'),
  )

  const chevron = element('span', '⌄', 'chevron')
  chevron.setAttribute('aria-hidden', 'true')
  summary.append(heading, chevron)
  details.append(summary)

  const hidden = active === 'Objectives'
    ? ['objective', 'target date', 'acceptance level', 'status', 'progress']
    : active === 'Meeting follow up'
      ? ['meeting date', 'topic discussed', 'solved?']
      : ['description', 'related goal', 'type', 'status', 'priority/impact']

  for (const [label, value] of Object.entries(entry.fields)) {
    if (hidden.some(prefix => normalize(label).startsWith(prefix))) {
      continue
    }

    addField(body, label, value)
  }


  body.append(
    element('small', `${active} · source row ${entry.sourceRow}`, 'source'),
  )

  details.append(body)
  return details
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
      element('p', 'No records. Import a customer spreadsheet to begin.', 'empty'),
    )
    return
  }

  const toolbar = element('div', '', 'toolbar')
  const overview = element('section', '', 'section-overview')
  const panel = element('section', '', 'data-panel')
  const panelHeading = element('div', '', 'panel-heading')
  const list = element(
    'div',
    '',
    active === 'Objectives' ? 'objective-grid' : 'record-list',
  )

  const search = document.createElement('input')
  search.type = 'search'
  search.placeholder = 'Search this section…'
  search.setAttribute('aria-label', 'Search this section')
  search.className = 'search'
  search.value = searchText

  if (active === 'Meeting follow up') {
    const dates = [...new Set(entries.map(meetingDate))]
      .sort((a, b) => b.localeCompare(a))

    // Start with the most recent date. Preserve subsequent selections.
    if (selectedDates === null) {
      selectedDates = new Set(dates)
    }

    if (selectedTopicStates === null) {
      selectedTopicStates = new Set(['pending', 'resolved', 'unknown'])
    }

    toolbar.append(
      createMultiSelect(
        'Meeting dates',
        dates.map((date): [string, string] => [date, formatDate(date)]),
        selectedDates,
        values => {
          selectedDates = values
          drawList()
        },
      ),
      createMultiSelect(
        'Topic status',
        [
          ['pending', 'Pending'],
          ['resolved', 'Resolved'],
          ['unknown', 'Not recorded'],
        ],
        selectedTopicStates,
        values => {
          selectedTopicStates = values
          drawList()
        },
      ),
    )
  }

  if (active === 'Tickets') {
    const types = [
      ...new Set(entries.map(entry => getField(entry, 'Type'))),
    ].sort()

    const statuses = [
      ...new Set(entries.map(entry => getField(entry, 'Status'))),
    ].sort()

    if (selectedTicketTypes === null) {
      selectedTicketTypes = new Set(types)
    }

    if (selectedTicketStatuses === null) {
      selectedTicketStatuses = new Set(statuses)
    }

    toolbar.append(
      createMultiSelect(
        'Ticket type',
        types.map((type): [string, string] => [
          type,
          type || 'Not specified',
        ]),
        selectedTicketTypes,
        values => {
          selectedTicketTypes = values
          drawList()
        },
      ),
      createMultiSelect(
        'Ticket status',
        statuses.map((status): [string, string] => [
          status,
          status || 'Not specified',
        ]),
        selectedTicketStatuses,
        values => {
          selectedTicketStatuses = values
          drawList()
        },
      ),
    )
  }

  toolbar.append(search)
  panel.append(panelHeading, list)
  content.append(overview, toolbar, panel)

  // Overview shortcuts start a fresh view so their counts match the results.
  function showTickets(status?: string) {
    selectedTicketTypes = null
    selectedTicketStatuses = status === undefined ? null : new Set([status])
    searchText = ''
    render()
    content.querySelector<HTMLElement>('.panel-heading')?.focus()
  }

  function showTopics(state?: string, date?: string) {
    selectedDates = date === undefined ? null : new Set([date])
    selectedTopicStates = state === undefined ? null : new Set([state])
    searchText = ''
    if (state !== undefined) {
      for (const entry of entries) {
        if (topicState(entry) === state &&
          (date === undefined || meetingDate(entry) === date)) {
          expandedMeetingDates.add(meetingDate(entry))
        }
      }
    }
    render()
    content.querySelector<HTMLElement>('.panel-heading')?.focus()
  }

  function drawOverview() {
    overview.replaceChildren()
    overview.hidden = active === 'Objectives'
    if (active === 'Objectives') return
    const grid = element('div', '', 'metrics-grid')
    const allSelected = (selected: Set<string> | null, values: string[]) =>
      values.every(value => selected?.has(value))

    if (active === 'Tickets') {
      overview.setAttribute('aria-label', 'Ticket overview')
      overview.append(element('h2', 'Ticket overview', 'overview-title'),
        element('p', 'All imported tickets. Select a number or a bar segment to explore a status.', 'overview-note'))
      const statuses = [...new Set(entries.map(entry => getField(entry, 'Status')))].sort()
      const types = [...new Set(entries.map(entry => getField(entry, 'Type')))]
      const unfiltered = !searchText && allSelected(selectedTicketTypes, types)
      const palettes: Record<string, string[]> = {
        success: ['#6fbd9d', '#99d8b0', '#57b9bb', '#b2ce80'],
        warning: ['#d8b166', '#e3c68a', '#dca67d', '#ecd7aa'],
        critical: ['#e67c7c', '#e49ead', '#cd91c2', '#f0ad91'],
        neutral: ['#82aad6', '#bb9ee8', '#69c5ce', '#aba6db'],
      }
      const usedColors: Record<string, number> = {}
      const parts = statuses.map(status => {
        const tone = statusTone(status)
        const palette = palettes[tone]!
        const index = usedColors[tone] || 0
        if (status) usedColors[tone] = index + 1
        return {
          label: status || 'No status',
          count: entries.filter(entry => getField(entry, 'Status') === status).length,
          color: status ? palette[index % palette.length]! : '#a6a6a6',
          onClick: () => showTickets(status),
        }
      })
      grid.append(metric('Total tickets', entries.length, 'View all tickets',
        () => showTickets(), unfiltered && allSelected(selectedTicketStatuses, statuses)))
      parts.forEach((part, index) => {
        grid.append(metric(part.label, part.count,
          `${Math.round(part.count / entries.length * 100)}% of all tickets · View tickets`,
          part.onClick, unfiltered && selectedTicketStatuses?.size === 1 &&
            selectedTicketStatuses.has(statuses[index]!), part.color))
      })
      overview.append(grid, distribution(parts, 'Tickets by status. Select a segment to filter.'))
    } else {
      overview.setAttribute('aria-label', 'Meeting overview')
      const counts = topicCounts(entries)
      const dates = [...new Set(entries.map(meetingDate))]
      const allDates = !searchText && allSelected(selectedDates, dates)
      overview.append(element('h2', 'Meeting overview', 'overview-title'),
        element('p', `Across all ${dates.length} meetings. Select a number to explore the topics.`, 'overview-note'))
      grid.append(
        metric('Topics discussed', counts.total, 'View all topics', () => showTopics(),
          allDates && allSelected(selectedTopicStates, ['resolved', 'pending', 'unknown'])),
        metric('Resolved', counts.resolved,
          `${Math.round(counts.resolved / counts.total * 100)}% of all topics`,
          () => showTopics('resolved'), allDates && selectedTopicStates?.size === 1 &&
            selectedTopicStates.has('resolved'), '#6fbd9d'),
        metric('Pending', counts.pending, 'Open pending topics →',
          () => showTopics('pending'), allDates && selectedTopicStates?.size === 1 &&
            selectedTopicStates.has('pending'), '#d8b166'),
      )
      if (counts.unknown) {
        grid.append(metric('Status not recorded', counts.unknown, 'Review these topics',
          () => showTopics('unknown'), allDates && selectedTopicStates?.size === 1 &&
            selectedTopicStates.has('unknown'), '#a6a6a6'))
      }
      overview.append(grid, distribution([
        { label: 'Resolved', count: counts.resolved, color: '#6fbd9d', onClick: () => showTopics('resolved') },
        { label: 'Pending', count: counts.pending, color: '#d8b166', onClick: () => showTopics('pending') },
        { label: 'Status not recorded', count: counts.unknown, color: '#a6a6a6', onClick: () => showTopics('unknown') },
      ], 'Meeting topics by resolution status'))
    }
  }

  function drawList() {
    list.replaceChildren()
    panelHeading.replaceChildren()
    panelHeading.tabIndex = -1
    drawOverview()

    let visible = [...entries]

    if (active === 'Objectives') {
      visible.sort((a, b) => objectiveRank(a) - objectiveRank(b))
    }

    if (active === 'Meeting follow up') {
      visible = visible.filter(entry =>
        selectedDates!.has(meetingDate(entry)) &&
        selectedTopicStates!.has(topicState(entry)),
      )

      panelHeading.append(
        element('h2', 'Meeting topics'),
        element('p', 'Topics are grouped by date. Open a topic to view details.'),
      )
    }

    if (active === 'Tickets') {
      visible = visible.filter(entry =>
        selectedTicketTypes!.has(getField(entry, 'Type')) &&
        selectedTicketStatuses!.has(getField(entry, 'Status')),
      )

      visible.sort((a, b) => Number(isFinished(a)) - Number(isFinished(b)))

      panelHeading.append(
        element('p', 'Open tickets first. Expand a ticket to view details.'),
      )
    }

    if (searchText.trim()) {
      const query = normalize(searchText)

      visible = visible.filter(entry => {
        const text = Object.entries(entry.fields)
          .filter(([key]) =>
            !(active === 'Tickets' &&
              normalize(key).startsWith('related goal')) &&
            !(active === 'Objectives' &&
              normalize(key).startsWith('target date')),
          )
          .map(([, value]) => value)
          .join(' ')

        return normalize(text).includes(query)
      })
    }

    const resultCount = element('p', `${visible.length} of ${entries.length} ${active === 'Meeting follow up' ? 'topics' : active.toLowerCase()} shown`, 'result-count')
    resultCount.setAttribute('role', 'status')
    panelHeading.append(resultCount)

    if (!visible.length) {
      list.append(
        element(
          'p',
          'No matching records. Check your selected options or search.',
          'empty',
        ),
      )
      return
    }

    if (active === 'Meeting follow up') {
      const groups = new Map<string, Entry[]>()

      for (const entry of visible) {
        const date = meetingDate(entry)
        const group = groups.get(date) || []
        group.push(entry)
        groups.set(date, group)
      }

      for (const date of [...groups.keys()].sort().reverse()) {
        const topics = groups.get(date)!
        const counts = topicCounts(entries.filter(entry => meetingDate(entry) === date))
        const wrapper = element('section', '', 'meeting-card')
        const group = document.createElement('details')
        group.className = 'meeting-group'
        group.open = expandedMeetingDates.has(date)

        const summary = document.createElement('summary')
        const heading = element('div', '', 'meeting-group-heading')

        heading.append(
          element('h2', formatDate(date)),
          element(
            'span',
            `${counts.total} discussed`,
            'meeting-topic-count',
          ),
        )
        heading.append(badge(`${counts.resolved} resolved`, 'success'),
          badge(`${counts.pending} pending`, 'warning'))
        if (counts.unknown) heading.append(badge(`${counts.unknown} not recorded`))
        if (topics.length !== counts.total) {
          heading.append(element('span', `${topics.length} shown`, 'meeting-topic-count'))
        }

        const arrow = element('span', '⌄', 'meeting-group-arrow')
        arrow.setAttribute('aria-hidden', 'true')
        summary.append(heading, arrow)

        const body = element('div', '', 'meeting-group-body')

        for (const entry of topics) {
          body.append(renderCard(entry))
        }

        group.addEventListener('toggle', () => {
          // Ignore toggle events from groups replaced by a filter update.
          if (!group.isConnected) return

          if (group.open) expandedMeetingDates.add(date)
          else expandedMeetingDates.delete(date)
        })

        group.append(summary, body)
        const footer = element('div', '', 'meeting-card-footer')
        footer.append(element('span', `${Math.round(counts.resolved / counts.total * 100)}% resolved`, 'meeting-resolution'))
        if (counts.pending) {
          const pending = document.createElement('button')
          pending.type = 'button'
          pending.className = 'pending-shortcut'
          pending.textContent = `View ${counts.pending} pending ${counts.pending === 1 ? 'topic' : 'topics'} →`
          pending.setAttribute('aria-label', `View ${counts.pending} pending topics for ${formatDate(date)}`)
          pending.onclick = () => showTopics('pending', date)
          footer.append(pending)
        }
        wrapper.append(group, footer)
        list.append(wrapper)
      }
    } else {
      for (const entry of visible) {
        list.append(renderCard(entry))
      }
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
    throw new Error(`The "${section}" worksheet was not found.`)
  }

  const sheet = workbook.Sheets[sheetName]!

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    blankrows: true,
    raw: true,
  })

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
    throw new Error(`Invalid headers in "${section}".`)
  }

  const result: Entry[] = []
  const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904)

  for (let index = 1; index < rows.length; index++) {
    const row = rows[index]!
    if (!String(row[requiredTitle] ?? '').trim()) continue

    const fields: Record<string, string> = {}

    headers.forEach((header, column) => {
      if (!header || /progress\s*%/i.test(header)) return

      const original = row[column]
      let value = String(original ?? '').trim()

      // Convert actual Excel date cells, preserving textual deadlines.
      const isDateField =
        /meeting date|target date|deadline/i.test(header)

      if (isDateField && typeof original === 'number') {
        const parsed = XLSX.SSF.parse_date_code(original, { date1904 })

        if (parsed) {
          value = [
            String(parsed.y).padStart(4, '0'),
            String(parsed.m).padStart(2, '0'),
            String(parsed.d).padStart(2, '0'),
          ].join('-')
        }
      }

      if (normalize(header) === 'solved?') {
        if (original === true || normalize(value) === 'true') value = 'Yes'
        if (original === false || normalize(value) === 'false') value = 'No'
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
  message.textContent = 'Reading spreadsheet…'

  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: 'array',
      cellDates: false,
    })

    const next = {
      Objectives: readSection(workbook, 'Objectives'),
      'Meeting follow up': readSection(workbook, 'Meeting follow up'),
      Tickets: readSection(workbook, 'Tickets'),
    }

    records = next
    expandedMeetingDates.clear()

    selectedDates = null
    selectedTopicStates = null
    selectedTicketTypes = null
    selectedTicketStatuses = null
    searchText = ''
    selectedDates = null
    selectedTopicStates = null
    selectedTicketTypes = null
    selectedTicketStatuses = null
    searchText = ''

    message.textContent =
      `${file.name} — ` +
      `${next.Objectives.length} objectives, ` +
      `${next['Meeting follow up'].length} meeting topics, ` +
      `${next.Tickets.length} tickets.`

    render()
  } catch (error) {
    message.textContent = error instanceof Error
      ? `Import failed: ${error.message}`
      : 'Unable to read this file.'
  } finally {
    input.disabled = false
    input.value = ''
  }
})

render()
