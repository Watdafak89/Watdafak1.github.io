document.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();

  document.querySelectorAll('.file-input').forEach((input) => {
    input.addEventListener('change', () => {
      const label = document.getElementById(input.dataset.label);
      label.textContent = input.files[0]?.name || '200MB per file · PDF';
      label.classList.toggle('selected-file', Boolean(input.files[0]));
    });
  });

  document.querySelectorAll('[data-step-target]').forEach((button) => {
    button.addEventListener('click', () => {
      const input = document.querySelector(`[name="${button.dataset.stepTarget}"]`);
      const nextValue = Math.max(Number(input.min || 0), Number(input.value) + Number(button.dataset.step));
      input.value = nextValue;
    });
  });

  document.getElementById('toggle-key').addEventListener('click', (event) => {
    const input = document.getElementById('api-key');
    const isHidden = input.type === 'password';
    input.type = isHidden ? 'text' : 'password';
    event.currentTarget.innerHTML = `<i data-lucide="${isHidden ? 'eye-off' : 'eye'}"></i>`;
    lucide.createIcons();
  });

  document.getElementById('course-form').addEventListener('submit', (event) => {
    event.preventDefault();
    const status = document.getElementById('status');
    const plan = document.getElementById('lesson-plan').files[0];
    status.textContent = plan
      ? `พร้อมวิเคราะห์ ${plan.name} · ตรวจสอบข้อมูลรายวิชาเรียบร้อยแล้ว`
      : 'กรุณาแนบไฟล์แผนการสอน PDF ก่อนเริ่มวิเคราะห์';
    status.style.color = plan ? '#168053' : '#c05c2e';
  });
});
