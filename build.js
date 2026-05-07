#!/usr/bin/env node
// Build standalone.html - single self-contained file, no CDN needed
const fs = require('fs');
const path = require('path');

let html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');

// 1. Remove Google Fonts (use system fonts instead)
html = html.replace(/<link rel="preconnect" href="https:\/\/fonts\.googleapis\.com">\n?/, '');
html = html.replace(/<link href="https:\/\/fonts\.googleapis\.com[^"]*" rel="stylesheet">\n?/, '');

// 2. Inline Chart.js
const chartJs = fs.readFileSync(path.join(__dirname, 'node_modules/chart.js/dist/chart.umd.js'), 'utf8');
html = html.replace(
  '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>',
  `<script>${chartJs}</script>`
);

// 3. Replace Font Awesome CDN with inline SVG icon system
const faIconsCss = `
<style id="fa-inline">
/* Inline SVG icon system - replaces Font Awesome */
.fas,.far,.fab { display:inline-block; vertical-align:-0.15em; width:1em; height:1em; }
.fas svg,.far svg,.fab svg { width:100%; height:100%; fill:currentColor; }
</style>
<script id="fa-inline-js">
(function(){
  const ICONS={
    'fa-diagram-project':'<svg viewBox="0 0 24 24"><rect x="3" y="3" width="5" height="5" rx="1"/><rect x="16" y="3" width="5" height="5" rx="1"/><rect x="9.5" y="16" width="5" height="5" rx="1"/><line x1="5.5" y1="8" x2="5.5" y2="12" stroke="currentColor" stroke-width="2"/><line x1="18.5" y1="8" x2="18.5" y2="12" stroke="currentColor" stroke-width="2"/><line x1="5.5" y1="12" x2="18.5" y2="12" stroke="currentColor" stroke-width="2"/><line x1="12" y1="12" x2="12" y2="16" stroke="currentColor" stroke-width="2"/></svg>',
    'fa-chart-pie':'<svg viewBox="0 0 24 24"><path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/></svg>',
    'fa-folder-open':'<svg viewBox="0 0 24 24"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/><polyline points="18 14 12 14 12 20"/></svg>',
    'fa-folder':'<svg viewBox="0 0 24 24"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    'fa-check-circle':'<svg viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
    'fa-check-double':'<svg viewBox="0 0 24 24"><polyline points="17 1 11 7 8 4"/><polyline points="21 5 15 11 12 8"/><polyline points="3 12 7 16 17 6"/></svg>',
    'fa-check':'<svg viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>',
    'fa-bell':'<svg viewBox="0 0 24 24"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>',
    'fa-bell-slash':'<svg viewBox="0 0 24 24"><path d="M13.73 21a2 2 0 0 1-3.46 0"/><path d="M18.63 13A17.89 17.89 0 0 1 18 8"/><path d="M6.26 6.26A5.86 5.86 0 0 0 6 8c0 7-3 9-3 9h14"/><path d="M18 8a6 6 0 0 0-9.33-5"/><line x1="1" y1="1" x2="23" y2="23"/></svg>',
    'fa-users':'<svg viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
    'fa-user':'<svg viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
    'fa-user-circle':'<svg viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
    'fa-user-plus':'<svg viewBox="0 0 24 24"><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="8.5" cy="7" r="4"/><line x1="20" y1="8" x2="20" y2="14"/><line x1="23" y1="11" x2="17" y2="11"/></svg>',
    'fa-user-slash':'<svg viewBox="0 0 24 24"><path d="M13.73 14A4 4 0 0 0 8.5 11.59"/><path d="M5.33 14A4 4 0 0 0 4 17v2"/><path d="M20 21v-2a4 4 0 0 0-.18-1.18"/><line x1="1" y1="1" x2="23" y2="23"/><circle cx="12" cy="7" r="4"/></svg>',
    'fa-bars':'<svg viewBox="0 0 24 24"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>',
    'fa-sign-out-alt':'<svg viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>',
    'fa-sign-in-alt':'<svg viewBox="0 0 24 24"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>',
    'fa-plus':'<svg viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>',
    'fa-times':'<svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>',
    'fa-edit':'<svg viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',
    'fa-trash':'<svg viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>',
    'fa-eye':'<svg viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>',
    'fa-search':'<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
    'fa-spinner':'<svg viewBox="0 0 24 24"><path d="M21 12a9 9 0 1 1-6.219-8.56" stroke-linecap="round"/></svg>',
    'fa-envelope':'<svg viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>',
    'fa-lock':'<svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>',
    'fa-info-circle':'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
    'fa-save':'<svg viewBox="0 0 24 24"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>',
    'fa-key':'<svg viewBox="0 0 24 24"><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></svg>',
    'fa-flask':'<svg viewBox="0 0 24 24"><path d="M9 3h6m-5 6 3 6m-5.5 0H17l-4.8-9.6A1 1 0 0 0 11.3 3H12m-3 6 1.5 3M5 21h14"/></svg>',
    'fa-fire':'<svg viewBox="0 0 24 24"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"/></svg>',
    'fa-wallet':'<svg viewBox="0 0 24 24"><path d="M21 12V7H5a2 2 0 0 1 0-4h14v4"/><path d="M3 5v14a2 2 0 0 0 2 2h16v-5"/><path d="M18 12a2 2 0 0 0 0 4h4v-4z"/></svg>',
    'fa-tasks':'<svg viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>',
    'fa-calendar-alt':'<svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
    'fa-calendar-check':'<svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><polyline points="9 16 11 18 15 14"/></svg>',
    'fa-clock':'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
    'fa-exclamation-triangle':'<svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    'fa-chart-bar':'<svg viewBox="0 0 24 24"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/><line x1="2" y1="20" x2="22" y2="20"/></svg>',
    'fa-chart-donut':'<svg viewBox="0 0 24 24"><path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/><circle cx="12" cy="12" r="4" fill="white"/></svg>',
    'fa-circle-half-stroke':'<svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 0 1 0 20z"/><circle cx="12" cy="12" r="10"/></svg>',
    'fa-receipt':'<svg viewBox="0 0 24 24"><polyline points="4 2 4 22 7.5 20 12 22 16.5 20 20 22 20 2"/><line x1="9" y1="9" x2="15" y2="9"/><line x1="9" y1="13" x2="15" y2="13"/></svg>',
    'fa-file-invoice-dollar':'<svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="12" x2="15" y2="12"/><line x1="9" y1="16" x2="15" y2="16"/><line x1="9" y1="8" x2="13" y2="8"/></svg>',
    'fa-th-large':'<svg viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>',
    'fa-list':'<svg viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>',
    'fa-chevron-right':'<svg viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg>',
    'fa-user-cog':'<svg viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><circle cx="19" cy="11" r="2"/><path d="M19 9v-.5M19 13v.5M17.2 10l-.4-.3M20.8 12l.4.3M17.2 12l-.4.3M20.8 10l.4-.3"/></svg>'
  };
  function applyIcons(root){
    (root||document).querySelectorAll('i[class]').forEach(function(el){
      const cls=Array.from(el.classList).find(function(c){return ICONS[c];});
      if(cls){
        const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
        svg.setAttribute('viewBox','0 0 24 24');
        svg.setAttribute('fill','none');
        svg.setAttribute('stroke','currentColor');
        svg.setAttribute('stroke-width','2');
        svg.setAttribute('stroke-linecap','round');
        svg.setAttribute('stroke-linejoin','round');
        svg.style.cssText='width:1em;height:1em;display:inline-block;vertical-align:-0.15em;flex-shrink:0';
        svg.innerHTML=ICONS[cls];
        // copy inner SVG content properly
        const tmp=document.createElement('div');
        tmp.innerHTML='<svg>'+ICONS[cls]+'</svg>';
        const innerSvg=tmp.querySelector('svg');
        if(innerSvg) Array.from(innerSvg.children).forEach(function(c){svg.appendChild(c.cloneNode(true));});
        el.innerHTML='';
        el.appendChild(svg);
        el.style.display='inline-flex';
        el.style.alignItems='center';
      }
    });
  }
  // Run after DOM ready and after dynamic content is added
  window.__applyIcons=applyIcons;
  document.addEventListener('DOMContentLoaded',function(){applyIcons();});
  // Observe DOM changes to apply icons to dynamically added content
  const obs=new MutationObserver(function(muts){muts.forEach(function(m){m.addedNodes.forEach(function(n){if(n.nodeType===1){applyIcons(n);}});});});
  document.addEventListener('DOMContentLoaded',function(){obs.observe(document.body,{childList:true,subtree:true});});
})();
</script>
`;

html = html.replace(
  '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">',
  faIconsCss
);

// 4. Fix font-family to use system fonts
html = html.replace("font-family:'Inter',sans-serif", "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif");

// 5. Add title indicator
html = html.replace('<title>Hệ thống Quản lý Dự án</title>', '<title>Quản lý Dự án - Standalone</title>');

fs.writeFileSync(path.join(__dirname, 'standalone.html'), html, 'utf8');
const size = fs.statSync(path.join(__dirname, 'standalone.html')).size;
console.log('✅ standalone.html created:', Math.round(size/1024), 'KB');
