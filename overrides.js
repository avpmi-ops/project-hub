// Dashboard panel overrides
function renderStageList(projects, stage, containerId, countId) {
  var stageProjects = projects
    .filter(function(p) { return p.Stage === stage; })
    .map(function(p) {
      var modDate = p.Modified ? new Date(p.Modified) : new Date();
      var days = Math.floor((Date.now() - modDate.getTime()) / (1000 * 60 * 60 * 24));
      return { title: p.Title, salesRep: p.Sales_Rep || '', days: days,
        color: days > 60 ? '#c0392b' : days > 30 ? '#d4a017' : '#2d8659' };
    })
    .sort(function(a, b) { return b.days - a.days; });
  if (countId) { var el = document.getElementById(countId); if (el) el.textContent = '(' + stageProjects.length + ')'; }
  var container = document.getElementById(containerId);
  if (!container) return;
  if (stageProjects.length === 0) {
    container.innerHTML = '<div style="color:#718096;font-size:13px;padding:20px 0;text-align:center;">No projects in this stage</div>';
    return;
  }
  var html = '<div class="health-grid">';
  stageProjects.forEach(function(item) {
    html += '<div class="health-item">';
    html += '<div class="dot" style="background:' + item.color + ';"></div>';
    html += '<div class="name">' + item.title;
    if (item.salesRep) html += '<span style="font-size:11px;color:#718096;font-weight:400;"> &middot; ' + item.salesRep + '</span>';
    html += '</div><div class="days">' + item.days + 'd</div></div>';
  });
  html += '</div>';
  container.innerHTML = html;
}
function renderPipeline(projects) {
  var panel = document.querySelector('#pipeline-chart') ? document.querySelector('#pipeline-chart').closest('.panel') : null;
  if (panel) { var h = panel.querySelector('h3'); if (h) h.innerHTML = 'Sales Handoff / Kickoff <span id="kickoff-count" style="font-size:12px;font-weight:500;color:#718096;"></span>'; }
  renderStageList(projects, 'Sales Handoff / Kickoff', 'pipeline-chart', 'kickoff-count');
}
function renderAlerts(projects) {
  var panel = document.querySelector('#alerts-panel') ? document.querySelector('#alerts-panel').closest('.panel') : null;
  if (panel) { var h = panel.querySelector('h3'); if (h) h.innerHTML = 'Go Live <span id="golive-count" style="font-size:12px;font-weight:500;color:#718096;"></span>'; }
  renderStageList(projects, 'Go Live', 'alerts-panel', 'golive-count');
}
function renderAging(projects) {
  var aging = projects.map(function(p) {
    var modDate = p.Modified ? new Date(p.Modified) : new Date();
    var days = Math.floor((Date.now() - modDate.getTime()) / (1000 * 60 * 60 * 24));
    return { title: p.Title, days: days, color: days > 120 ? '#c0392b' : days > 60 ? '#d4a017' : '#2d8659' };
  }).sort(function(a, b) { return b.days - a.days; }).slice(0, 12);
  var html = '<div class="health-grid">';
  aging.forEach(function(item) {
    html += '<div class="health-item"><div class="dot" style="background:' + item.color + ';"></div><div class="name">' + item.title + '</div><div class="days">' + item.days + 'd</div></div>';
  });
  html += '</div>';
  var el = document.getElementById('aging-panel');
  if (el) el.innerHTML = html;
}