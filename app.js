(() => {
  const API = String(window.CDE_CONFIG?.apiUrl || '');
  const reasons = ['Enfermedad','Cita médica','Viaje','Motivo familiar','Ausencia justificada','Ausencia no justificada','Otro'];
  const $ = selector => document.querySelector(selector);
  let session = null;
  let students = [];
  let pendingData = null;

  function configured() { return API.startsWith('https://script.google.com/') && API.endsWith('/exec'); }
  async function api(action, data = {}) {
    if (!configured()) throw new Error('La URL de Apps Script todavía no está configurada en config.js.');
    const response = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action, ...data }) });
    const result = await response.json();
    if (!result.ok && result.code !== 'DUPLICATE') throw new Error(result.message || 'No fue posible completar la solicitud.');
    return result;
  }
  function today() { return new Intl.DateTimeFormat('en-CA',{timeZone:'America/Costa_Rica',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()); }
  function show(id) { ['login-view','attendance-view','success-view'].forEach(view => $('#'+view).classList.toggle('hidden',view!==id)); $('#logout').classList.toggle('hidden',id==='login-view'); }
  function setBusy(button, busy, busyText) { if (!button.dataset.label) button.dataset.label = button.textContent; button.disabled = busy; button.textContent = busy ? busyText : button.dataset.label; }
  function escapeHtml(value) { return String(value || '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char])); }

  function renderStudents() {
    $('#student-list').innerHTML = students.map((student,index) => `<article class="student-card" data-index="${index}"><div class="student-info"><b>${escapeHtml(student.name)}</b><small>${escapeHtml(student.id)}${student.provisional?' · Provisional':''}</small></div><div class="attendance-choice"><label><input type="radio" name="status-${index}" value="P">✓ Presente</label><label><input type="radio" name="status-${index}" value="T">◷ Tardía</label><label><input type="radio" name="status-${index}" value="A">— Ausente</label></div><div class="absence-fields hidden"><select aria-label="Motivo de ausencia"><option value="">Selecciona el motivo</option>${reasons.map(reason=>`<option>${reason}</option>`).join('')}</select><input class="other-detail hidden" placeholder="Especifica el motivo" aria-label="Detalle de otro motivo"></div></article>`).join('');
    updateProgress();
  }
  function updateProgress() {
    let present=0, late=0, absent=0, valid=0;
    document.querySelectorAll('.student-card').forEach(card => {
      const status=card.querySelector('input:checked')?.value, reason=card.querySelector('select').value, detail=card.querySelector('.other-detail').value.trim();
      if(status==='P'){present++;valid++;}
      if(status==='T'){late++;valid++;}
      if(status==='A'){absent++;if(reason&&(reason!=='Otro'||detail)) valid++;}
    });
    const total=students.length, pct=total?Math.round(valid/total*100):0;
    $('#progress-text').textContent=`${valid} de ${total} estudiantes registrados`;
    $('#progress-percent').textContent=`${pct}%`; $('#progress-bar').style.width=`${pct}%`;
    $('#present-count').textContent=present; $('#late-count').textContent=late; $('#absent-count').textContent=absent;
    $('#submit-summary').textContent=valid===total?`${present} presentes · ${late} tardías · ${absent} ausentes`:'Completa la lista';
    $('#submit-attendance').disabled=valid!==total;
    return {present,late,absent,valid,total};
  }
  function buildPayload() {
    return { token:session.token, date:$('#attendance-date').value, records:[...document.querySelectorAll('.student-card')].map((card,index)=>{const selected=card.querySelector('input:checked').value, reason=card.querySelector('select').value, statusCode=selected==='A'&&reason==='Ausencia justificada'?'J':selected;return{student_id:students[index].id,student_name:students[index].name,status_code:statusCode,absence_reason:selected==='A'?reason:'',other_detail:reason==='Otro'?card.querySelector('.other-detail').value.trim():''};}) };
  }
  async function refreshCalendar() {
    $('#form-error').textContent='';
    try { const result=await api('check',{token:session.token,date:$('#attendance-date').value}); const calendar=result.calendar; $('#school-year').textContent=`Año lectivo ${calendar.schoolYear}`; $('#block-week').textContent=`${calendar.block} · ${calendar.week}`; $('#calendar-note').textContent=calendar.month; }
    catch(error){ $('#school-year').textContent='Fecha no configurada'; $('#block-week').textContent='Revisa Attendance_Calendar'; $('#calendar-note').textContent=''; $('#form-error').textContent=error.message; }
  }
  function renderSuccess(result) {
    const date=new Date(`${pendingData.date}T12:00:00`).toLocaleDateString('es-CR',{dateStyle:'long'});
    $('#success-details').innerHTML=[['Fecha',date],['Grupo',session.teacher.group],['Profesora',session.teacher.name],['Total',result.total],['Presentes',result.present],['Tardías',result.late],['Ausentes',result.absent],['Justificadas',result.justified]].map(([label,value])=>`<div class="receipt-row"><span>${label}</span><b>${value}</b></div>`).join('');
    $('#notification-status').textContent='✓ Reporte enviado por correo a Sandra y a la profesora.';
    show('success-view');
  }
  async function send(confirmUpdate=false) {
    const button=$('#submit-attendance'); setBusy(button,true,confirmUpdate?'Actualizando…':'Enviando…'); $('#form-error').textContent='';
    try { const result=await api('submit',{...pendingData,confirmUpdate}); if(result.code==='DUPLICATE'){setBusy(button,false);$('#update-dialog').showModal();return;} renderSuccess(result); }
    catch(error){ $('#form-error').textContent=error.message; setBusy(button,false); }
  }

  $('#login-form').addEventListener('submit',async event=>{event.preventDefault();const button=$('#login-button');setBusy(button,true,'Validando…');$('#login-error').textContent='';try{const result=await api('login',{code:$('#access-code').value});session={token:result.token,teacher:result.teacher};students=result.students;$('#group-name').textContent=result.teacher.group;$('#teacher-name').textContent=`Profesora ${result.teacher.name}`;const warning=$('#code-warning');warning.classList.toggle('hidden',!result.codeWarnings?.length);warning.textContent=result.codeWarnings?.length?`Aviso: ${result.codeWarnings.length} código(s) provisional(es) en este grupo. La asistencia puede registrarse, pero Administración debe completar el año de ingreso para confirmarlos.`:'';$('#attendance-date').value=today();renderStudents();show('attendance-view');await refreshCalendar();}catch(error){$('#login-error').textContent=error.message;}finally{setBusy(button,false);}});
  $('#toggle-code').addEventListener('click',()=>{const input=$('#access-code');input.type=input.type==='password'?'text':'password';$('#toggle-code').textContent=input.type==='password'?'Ver':'Ocultar';});
  $('#attendance-date').addEventListener('change',refreshCalendar);
  $('#student-list').addEventListener('change',event=>{const card=event.target.closest('.student-card');if(!card)return;const status=card.querySelector('input:checked')?.value, fields=card.querySelector('.absence-fields'), other=card.querySelector('.other-detail');fields.classList.toggle('hidden',status!=='A');card.classList.toggle('is-absent',status==='A');other.classList.toggle('hidden',card.querySelector('select').value!=='Otro');updateProgress();});
  $('#student-list').addEventListener('input',updateProgress);
  $('#mark-all').addEventListener('click',()=>{document.querySelectorAll('.student-card input[value="P"]').forEach(input=>input.checked=true);document.querySelectorAll('.student-card').forEach(card=>{card.classList.remove('is-absent');card.querySelector('.absence-fields').classList.add('hidden');});updateProgress();});
  $('#attendance-form').addEventListener('submit',event=>{event.preventDefault();if(updateProgress().valid!==students.length){$('#form-error').textContent='Completa la asistencia y los motivos pendientes.';return;}pendingData=buildPayload();send(false);});
  $('#update-dialog').addEventListener('close',()=>{if($('#update-dialog').returnValue==='confirm')send(true);});
  $('#logout').addEventListener('click',()=>{session=null;students=[];$('#access-code').value='';show('login-view');});
  $('#back-home').addEventListener('click',()=>$('#logout').click());
  if(!configured()) $('#login-error').textContent='Pendiente: configura la URL de Apps Script en config.js.';
})();
