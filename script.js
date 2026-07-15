
const menuBtn=document.querySelector('.menu-btn');const nav=document.querySelector('.nav');
menuBtn?.addEventListener('click',()=>{const open=nav.classList.toggle('open');menuBtn.setAttribute('aria-expanded',String(open));});
document.querySelectorAll('.nav a').forEach(a=>a.addEventListener('click',()=>nav.classList.remove('open')));
const observer=new IntersectionObserver(entries=>entries.forEach(e=>{if(e.isIntersecting)e.target.classList.add('visible')}),{threshold:.12});
document.querySelectorAll('.reveal').forEach(el=>observer.observe(el));
document.querySelectorAll('[data-open-modal]').forEach(btn=>btn.addEventListener('click',()=>document.getElementById(btn.dataset.openModal)?.showModal()));
document.querySelectorAll('dialog').forEach(modal=>{
  modal.querySelector('.modal-close')?.addEventListener('click',()=>modal.close());
  modal.addEventListener('click',e=>{const box=modal.getBoundingClientRect();const inside=e.clientX>=box.left&&e.clientX<=box.right&&e.clientY>=box.top&&e.clientY<=box.bottom;if(!inside)modal.close();});
});
