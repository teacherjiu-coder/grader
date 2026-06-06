(function () {
  'use strict';

  const AREAS = ['기본작업', '계산작업', '분석작업', '기타작업'];
  const ALL_ROUNDS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

  let state = {
    level: null,
    round: null,
    keyTitle: '',
  };

  const $ = (sel) => document.querySelector(sel);

  const stepLevel = $('#step-level');
  const stepRound = $('#step-round');
  const stepUpload = $('#step-upload');
  const stepResult = $('#step-result');
  const roundsEl = $('#rounds');
  const roundHeading = $('#round-heading');
  const uploadTag = $('#upload-tag');
  const dropZone = $('#drop-zone');
  const fileInput = $('#file-input');
  const loading = $('#loading');
  const toast = $('#toast');

  function showStep(name) {
    stepLevel.classList.toggle('hide', name !== 'level');
    stepRound.classList.toggle('hide', name !== 'round');
    stepUpload.classList.toggle('hide', name !== 'upload');
    stepResult.classList.toggle('hide', name !== 'result');
  }

  function setLoading(on) {
    loading.classList.toggle('hide', !on);
  }

  function showToast(msg, ms = 5000) {
    toast.textContent = msg;
    toast.classList.remove('hide');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toast.classList.add('hide'), ms);
  }

  function isRoundEnabled(level, round) {
    if (level === '2급') return round >= 1 && round <= 10;
    if (level === '1급') return round === 1;
    return false;
  }

  function selectLevel(level) {
    state.level = level;
    state.round = null;
    roundHeading.textContent = `컴활 ${level} · 회차 선택`;
    roundsEl.innerHTML = '';

    ALL_ROUNDS.forEach((n) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'round-btn';
      const enabled = isRoundEnabled(level, n);
      btn.disabled = !enabled;
      btn.innerHTML = enabled
        ? `${n}회`
        : `${n}회<span class="soon">준비 중</span>`;
      if (enabled) btn.onclick = () => selectRound(n);
      roundsEl.appendChild(btn);
    });

    showStep('round');
  }

  async function selectRound(round) {
    state.round = round;
    try {
      const res = await fetch(`answers/${state.level}/${round}.json`);
      if (!res.ok) throw new Error('정답키 없음');
      const key = await res.json();
      state.keyTitle = key.title || `컴활 ${state.level} ${round}회`;
    } catch {
      showToast('정답키를 불러올 수 없어요. 회차를 다시 선택해주세요.');
      return;
    }
    uploadTag.textContent = state.keyTitle;
    showStep('upload');
  }

  async function gradeFile(file) {
    const buf = new Uint8Array(await file.arrayBuffer());
    const res = await fetch(`answers/${state.level}/${state.round}.json`);
    if (!res.ok) throw new Error('정답키를 불러올 수 없습니다.');
    const KEY = await res.json();
    const result = await window.gradeWorkbook(buf, { XLSX, JSZip, DOMParser }, KEY);
    console.log('[gradeWorkbook result]', result);
    return result;
  }

  function isFormatError(err) {
    const m = String(err && err.message ? err.message : err);
    return /sheet|workbook|Cannot read|undefined|null|zip|xml|parse/i.test(m);
  }

  async function handleFile(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsm', 'xlsx'].includes(ext)) {
      showToast('.xlsm 또는 .xlsx 파일만 올릴 수 있어요.');
      return;
    }

    setLoading(true);
    try {
      const result = await gradeFile(file);
      renderResult(result);
      showStep('result');
      stepResult.scrollIntoView({ behavior: 'smooth' });
    } catch (err) {
      console.error(err);
      if (isFormatError(err)) {
        showToast(
          `이 파일은 ${state.level} ${state.round}회 기출 양식이 아닌 것 같아요. 회차를 확인해주세요.`
        );
      } else {
        showToast('채점 중 오류: ' + (err.message || '알 수 없는 오류'));
      }
    } finally {
      setLoading(false);
      fileInput.value = '';
    }
  }

  function renderResult(r) {
    $('#score-num').innerHTML =
      `${r.total}<span class="unit">점</span>`;
    $('#pass-badge').innerHTML = r.pass
      ? '<span class="pass-badge pass">합격 예상 🎉</span>'
      : '<span class="pass-badge fail">조금만 더! (70점 미만)</span>';

    const grid = $('#area-grid');
    grid.innerHTML = '';
    AREAS.forEach((area) => {
      const rows = r.results.filter((x) => x.area === area);
      if (!rows.length) return;
      const earned = rows.reduce((s, x) => s + x.earned, 0);
      const max = rows.reduce((s, x) => s + x.points, 0);
      const pct = max ? Math.round((earned / max) * 100) : 0;
      const card = document.createElement('div');
      card.className = 'area-card';
      card.innerHTML =
        `<div class="name">${area}</div>` +
        `<div class="pts">${earned}<span> / ${max}점</span></div>` +
        `<div class="progress"><div class="progress-bar" style="width:${pct}%"></div></div>`;
      grid.appendChild(card);
    });

    const list = $('#problem-list');
    list.innerHTML = '';
    AREAS.forEach((area) => {
      const rows = r.results.filter((x) => x.area === area);
      if (!rows.length) return;
      const earned = rows.reduce((s, x) => s + x.earned, 0);
      const max = rows.reduce((s, x) => s + x.points, 0);
      const group = document.createElement('div');
      group.className = 'area-group';
      group.innerHTML =
        `<div class="area-group-head">` +
        `<span class="an">${area}</span>` +
        `<span class="ap">${earned} / ${max}점</span></div>`;

      rows.forEach((x) => {
        const row = document.createElement('div');
        row.className = 'prow';
        let cls, sym, body;

        if (x.manual) {
          cls = 'man';
          sym = '👀';
          body =
            `<div class="pname">${esc(x.id)}</div>` +
            `<div class="pman">확인 필요 (만점 인정)</div>`;
        } else if (x.ok) {
          cls = 'ok';
          sym = '✓';
          body = `<div class="pname">${esc(x.id)}</div>`;
        } else {
          cls = 'no';
          sym = '✕';
          body =
            `<div class="pname">${esc(x.id)}</div>` +
            (x.msg ? `<div class="pmsg">${esc(x.msg)}</div>` : '') +
            (x.answer ? `<div class="pans"><b>정답:</b> ${esc(x.answer)}</div>` : '');
        }

        row.innerHTML =
          `<div class="ico ${cls}">${sym}</div>` +
          `<div class="pbody">${body}</div>` +
          `<div class="ppts">${x.earned}<small>/${x.points}</small></div>`;
        group.appendChild(row);
      });
      list.appendChild(group);
    });
  }

  function esc(s) {
    const d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
  }

  function resetAll() {
    state = { level: null, round: null, keyTitle: '' };
    fileInput.value = '';
    showStep('level');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  document.querySelectorAll('.lvbtn').forEach((btn) => {
    btn.addEventListener('click', () => selectLevel(btn.dataset.level));
  });

  $('#back-to-level').addEventListener('click', () => {
    state.level = null;
    showStep('level');
  });

  $('#back-to-round').addEventListener('click', () => {
    state.round = null;
    showStep('round');
  });

  $('#retry-btn').addEventListener('click', resetAll);

  $('#pick-file').addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('click', (e) => {
    if (e.target.id !== 'pick-file') fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('hover');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('hover'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('hover');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
})();
