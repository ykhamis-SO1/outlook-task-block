// =============================================
// FILE: taskpane/taskpane.js
// =============================================
// Save as ./taskpane/taskpane.js

`(function(){
  /** Utility **/
  const $ = (sel) => document.querySelector(sel);
  const tasksEl = $('#tasks');
  const totalPlannedEl = $('#totalPlanned');
  const overUnderEl = $('#overUnder');

  function nowRoundedTo(minutes) {
    const d = new Date();
    const ms = 1000 * 60 * minutes;
    return new Date(Math.ceil(d.getTime() / ms) * ms);
  }

  function renderTaskRow(id, title = '', mins = 15) {
    const row = document.createElement('div');
    row.className = 'task';
    row.dataset.id = id;
    row.innerHTML = `
      <input type="text" placeholder="Task name" value="${title}" />
      <input type="number" min="5" step="5" value="${mins}" />
      <div class="del" title="Remove">×</div>
    `;
    row.querySelector('.del').onclick = () => { row.remove(); recalc(); };
    row.querySelectorAll('input').forEach(inp => inp.addEventListener('input', recalc));
    tasksEl.appendChild(row);
  }

  function getTasks() {
    return Array.from(document.querySelectorAll('.task')).map(row => {
      const [titleEl, minsEl] = row.querySelectorAll('input');
      return { title: titleEl.value.trim() || 'Task', mins: Math.max(5, parseInt(minsEl.value || '0', 10)) };
    });
  }

  function recalc() {
    const tasks = getTasks();
    const total = tasks.reduce((s,t)=> s + (t.mins||0), 0);
    totalPlannedEl.textContent = total;
    const block = parseInt($('#blockMinutes').value || '60', 10);
    const delta = block - total;
    overUnderEl.textContent = delta === 0 ? 'Perfect fit' : (delta > 0 ? `${delta} min free` : `${-delta} min over`);
    overUnderEl.className = delta < 0 ? 'over' : (delta > 0 ? 'under' : '');
  }

  function scheduleFrom(startDate, tasks) {
    let t = new Date(startDate);
    return tasks.map(task => {
      const st = new Date(t);
      const et = new Date(t.getTime() + task.mins*60000);
      t = et;
      return { ...task, start: st, end: et };
    });
  }

  function fmtTime(d) {
    return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  function buildTable(sched) {
    let html = '<table class="table"><thead><tr><th>Time</th><th>Task</th><th>Minutes</th></tr></thead><tbody>';
    for (const s of sched) {
      html += `<tr><td>${fmtTime(s.start)}–${fmtTime(s.end)}</td><td>${escapeHtml(s.title)}</td><td>${s.mins}</td></tr>`;
    }
    html += '</tbody></table>';
    return html;
  }

  function buildChecklist(sched) {
    return '<div>' + sched.map(s => `□ ${fmtTime(s.start)}–${fmtTime(s.end)} — ${escapeHtml(s.title)} (${s.mins}m)`).join('<br/>') + '</div>';
  }

  function escapeHtml(s) {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  async function getOrDefaultStart() {
    return new Promise((resolve) => {
      const item = Office.context.mailbox.item;
      if (item?.start && item?.end && item.start.getAsync) {
        item.start.getAsync(r => {
          if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
            resolve(new Date(r.value));
          } else {
            resolve(nowRoundedTo(5));
          }
        });
      } else {
        resolve(nowRoundedTo(5));
      }
    });
  }

  async function setEventTimes(start, minutes) {
    const item = Office.context.mailbox.item;
    return new Promise((resolve) => {
      if (!(item?.start?.setAsync && item?.end?.setAsync)) return resolve(false);
      const end = new Date(start.getTime() + minutes*60000);
      item.start.setAsync(start, () => {
        item.end.setAsync(end, () => resolve(true));
      });
    });
  }

  async function insertIntoBody(html, subjectPrefix) {
    return new Promise((resolve) => {
      const item = Office.context.mailbox.item;
      if (item?.body?.setAsync) {
        // Prepend content to body
        item.body.getAsync(Office.CoercionType.Html, r => {
          const existing = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? r.value : '';
          const composed = `<div><strong>Task Plan</strong></div>${html}<hr/>` + existing;
          item.body.setAsync(composed, { coercionType: Office.CoercionType.Html }, () => resolve(true));
        });
      } else {
        resolve(false);
      }

      if (item?.subject?.setAsync) {
        item.subject.getAsync(r => {
          const cur = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? r.value : '';
          if (!cur.startsWith(subjectPrefix)) {
            item.subject.setAsync(subjectPrefix + cur);
          }
        });
      }
    });
  }

  function refreshPreview() {
    const tasks = getTasks();
    const start = new Date($('#startTime').value || nowRoundedTo(5));
    const sched = scheduleFrom(start, tasks);
    const html = $('#formatStyle').value === 'table' ? buildTable(sched) : buildChecklist(sched);
    $('#preview').innerHTML = html;
  }

  // Wire-up UI
  function addTaskDefaults() {
    renderTaskRow(1, 'Deep work', 25);
    renderTaskRow(2, 'Admin/Email', 10);
    renderTaskRow(3, 'Ops follow-up', 15);
    recalc();
  }

  function initUi() {
    $('#btnAddTask').onclick = () => { renderTaskRow(Date.now()); recalc(); };
    $('#btnPreview').onclick = refreshPreview;
    $('#blockMinutes').addEventListener('input', recalc);
    $('#formatStyle').addEventListener('change', refreshPreview);

    $('#btnUseEventTimes').onclick = async () => {
      const item = Office.context.mailbox.item;
      if (item?.start?.getAsync) {
        item.start.getAsync(r => {
          const st = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? new Date(r.value) : nowRoundedTo(5);
          $('#startTime').value = toLocalInput(st);
          refreshPreview();
        });
      }
    };

    $('#btnSetEventTimes').onclick = async () => {
      const minutes = parseInt($('#blockMinutes').value || '60', 10);
      const start = new Date($('#startTime').value || nowRoundedTo(5));
      await setEventTimes(start, minutes);
    };

    $('#btnInsert').onclick = async () => {
      const minutes = parseInt($('#blockMinutes').value || '60', 10);
      const start = new Date($('#startTime').value || (await getOrDefaultStart()));
      const tasks = getTasks();
      const sched = scheduleFrom(start, tasks);
      const html = $('#formatStyle').value === 'table' ? buildTable(sched) : buildChecklist(sched);
      const prefix = $('#subjectPrefix').value || '';
      await insertIntoBody(html, prefix);
      recalc();
      refreshPreview();
    };
  }

  function toLocalInput(dt) {
    const pad = (n) => String(n).padStart(2,'0');
    const y = dt.getFullYear();
    const m = pad(dt.getMonth()+1);
    const d = pad(dt.getDate());
    const hh = pad(dt.getHours());
    const mm = pad(dt.getMinutes());
    return `${y}-${m}-${d}T${hh}:${mm}`;
  }

  Office.onReady(() => {
    const st = nowRoundedTo(5);
    $('#startTime').value = toLocalInput(st);
    addTaskDefaults();
    initUi();
  });
})();`
