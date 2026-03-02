(function () {
  function ensureMaterialSymbolsLoaded() {
    if (document.getElementById("unified-nav-material-symbols")) return;
    const link = document.createElement("link");
    link.id = "unified-nav-material-symbols";
    link.rel = "stylesheet";
    link.href = "/material-symbols.css";
    document.head.appendChild(link);
  }

  function resolveReportHref(fileName) {
    if (window.location.protocol === "file:") return "report_html/" + fileName;
    return "/" + fileName;
  }

  const NAV_SECTIONS = [
    {
      id: "reports",
      mainLabel: "Reports",
      mainIcon: "description",
      title: "Reports",
      items: [
        { title: "Dashboard", href: resolveReportHref("dashboard.html"), icon: "space_dashboard", file: "dashboard.html" },
        { title: "Nested View Report", href: resolveReportHref("nested_view_report.html"), icon: "account_tree", file: "nested_view_report.html" },
        { title: "Missed Entries Report", href: resolveReportHref("missed_entries.html"), icon: "event_busy", file: "missed_entries.html" },
        { title: "Assignee Hours Report", href: resolveReportHref("assignee_hours_report.html"), icon: "schedule", file: "assignee_hours_report.html" },
        { title: "Employee Performance", href: resolveReportHref("employee_performance_report.html"), icon: "monitoring", file: "employee_performance_report.html" },
        { title: "RLT Leave Report", href: resolveReportHref("rlt_leave_report.html"), icon: "beach_access", file: "rlt_leave_report.html" },
        { title: "Leaves Planned Calendar", href: resolveReportHref("leaves_planned_calendar.html"), icon: "calendar_month", file: "leaves_planned_calendar.html" },
        { title: "RnD Data Story", href: resolveReportHref("rnd_data_story.html"), icon: "auto_stories", file: "rnd_data_story.html" },
        { title: "Planned RMIs", href: resolveReportHref("planned_rmis_report.html"), icon: "assignment_turned_in", file: "planned_rmis_report.html" },
        { title: "Phase RMI Gantt", href: resolveReportHref("phase_rmi_gantt_report.html"), icon: "view_timeline", file: "phase_rmi_gantt_report.html" },
        { title: "IPP Meeting Dashboard", href: resolveReportHref("ipp_meeting_dashboard.html"), icon: "groups", file: "ipp_meeting_dashboard.html" }
      ]
    },
    {
      id: "admin-settings",
      mainLabel: "Admin Settings",
      mainIcon: "admin_panel_settings",
      title: "Admin Settings",
      items: [
        { title: "Capacity Settings", href: "/settings/capacity", icon: "tune", path: "/settings/capacity" },
        { title: "Performance Settings", href: "/settings/performance", icon: "speed", path: "/settings/performance" },
        { title: "Report Entities", href: "/settings/report-entities", icon: "dataset", path: "/settings/report-entities" },
        { title: "Manage Fields", href: "/settings/manage-fields", icon: "list_alt", path: "/settings/manage-fields" },
        { title: "Projects", href: "/settings/projects", icon: "work", path: "/settings/projects" },
        { title: "Epic Dropdowns", href: "/settings/epics-dropdown-options", icon: "arrow_drop_down_circle", path: "/settings/epics-dropdown-options" },
        { title: "Epic Phases", href: "/settings/epic-phases", icon: "alt_route", path: "/settings/epic-phases" },
        { title: "Epics Planner", href: "/settings/epics-management", icon: "event_note", path: "/settings/epics-management" }
      ]
    }
  ];

  const storageKey = "unified-report-nav-collapsed";
  const isMobile = window.matchMedia("(max-width: 959px)").matches;
  const currentPath = String(window.location.pathname || "").replace(/\\/g, "/").toLowerCase();
  const currentFile = (window.location.pathname.split("/").pop() || "").toLowerCase();
  ensureMaterialSymbolsLoaded();

  const nav = document.createElement("aside");
  nav.className = "unified-nav";
  nav.setAttribute("aria-label", "Reports navigation");

  const savedCollapsed = localStorage.getItem(storageKey) === "1";
  if (!isMobile && savedCollapsed) {
    nav.classList.add("is-collapsed");
    document.body.classList.add("unified-nav-collapsed");
  }

  const head = document.createElement("div");
  head.className = "unified-nav-head";

  const brand = document.createElement("div");
  brand.innerHTML = '<h2 class="unified-nav-title">Reports</h2><div class="unified-nav-sub">Unified Navigation</div>';

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "unified-nav-toggle";
  toggle.textContent = "\u2261";
  toggle.setAttribute("aria-label", "Toggle navigation");

  head.appendChild(brand);
  head.appendChild(toggle);
  nav.appendChild(head);

  const body = document.createElement("div");
  body.className = "unified-nav-body";

  const mainNav = document.createElement("nav");
  mainNav.className = "unified-nav-main";
  mainNav.setAttribute("aria-label", "Main navigation");

  const list = document.createElement("nav");
  list.className = "unified-nav-list";
  list.setAttribute("aria-label", "Secondary navigation");

  const activeSection = NAV_SECTIONS.find((section) =>
    section.items.some((item) => {
      const isActiveFile = item.file && currentFile === item.file.toLowerCase();
      const isActivePath = item.path && currentPath.indexOf(item.path.toLowerCase()) === 0;
      return Boolean(isActiveFile || isActivePath);
    })
  );
  let activeSectionId = activeSection ? activeSection.id : NAV_SECTIONS[0].id;
  const reportsSection = NAV_SECTIONS.find((section) => section.id === "reports") || NAV_SECTIONS[0];
  const adminSection = NAV_SECTIONS.find((section) => section.id === "admin-settings");
  let isAdminExpanded = activeSectionId === "admin-settings";
  const mainButtonsBySectionId = {};

  function renderMainNav() {
    reportsSection.items.forEach((item) => {
      const link = document.createElement("a");
      link.className = "unified-nav-main-link";
      link.href = item.href;
      const isActiveFile = item.file && currentFile === item.file.toLowerCase();
      const isActivePath = item.path && currentPath.indexOf(item.path.toLowerCase()) === 0;
      if (isActiveFile || isActivePath) {
        link.classList.add("is-active");
        link.setAttribute("aria-current", "page");
      }
      link.addEventListener("click", function () {
        isAdminExpanded = false;
        renderSecondaryNav();
      });

      const icon = document.createElement("span");
      icon.className = "unified-nav-main-icon material-symbols-outlined";
      icon.setAttribute("aria-hidden", "true");
      icon.textContent = item.icon;

      const label = document.createElement("span");
      label.className = "unified-nav-main-label";
      label.textContent = item.title;

      link.appendChild(icon);
      link.appendChild(label);
      mainNav.appendChild(link);
    });

    if (adminSection) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "unified-nav-main-link";
      button.dataset.sectionId = adminSection.id;

      const icon = document.createElement("span");
      icon.className = "unified-nav-main-icon material-symbols-outlined";
      icon.setAttribute("aria-hidden", "true");
      icon.textContent = adminSection.mainIcon;

      const label = document.createElement("span");
      label.className = "unified-nav-main-label";
      label.textContent = adminSection.mainLabel;

      button.appendChild(icon);
      button.appendChild(label);
      button.addEventListener("click", function () {
        isAdminExpanded = true;
        renderSecondaryNav();
      });
      mainButtonsBySectionId[adminSection.id] = button;
      mainNav.appendChild(button);
    }
  }

  function renderSecondaryNav() {
    list.innerHTML = "";
    const adminBtn = adminSection ? mainButtonsBySectionId[adminSection.id] : null;
    if (adminBtn) adminBtn.classList.toggle("is-active", isAdminExpanded);
    if (!adminSection || !isAdminExpanded) {
      list.classList.add("is-hidden");
      return;
    }
    list.classList.remove("is-hidden");

    const sectionEl = document.createElement("section");
    sectionEl.className = "unified-nav-section";

    const sectionTitle = document.createElement("h3");
    sectionTitle.className = "unified-nav-section-title";
    sectionTitle.textContent = adminSection.title;
    sectionEl.appendChild(sectionTitle);

    adminSection.items.forEach((item) => {
      const link = document.createElement("a");
      link.className = "unified-nav-link";
      link.href = item.href;
      const isActiveFile = item.file && currentFile === item.file.toLowerCase();
      const isActivePath = item.path && currentPath.indexOf(item.path.toLowerCase()) === 0;
      if (isActiveFile || isActivePath) {
        link.classList.add("is-active");
        link.setAttribute("aria-current", "page");
      }

      const icon = document.createElement("span");
      icon.className = "unified-nav-icon material-symbols-outlined";
      icon.setAttribute("aria-hidden", "true");
      icon.textContent = item.icon;

      const label = document.createElement("span");
      label.className = "unified-nav-label";
      label.textContent = item.title;

      link.appendChild(icon);
      link.appendChild(label);
      sectionEl.appendChild(link);
    });
    list.appendChild(sectionEl);
  }

  renderMainNav();
  renderSecondaryNav();
  body.appendChild(mainNav);
  body.appendChild(list);
  nav.appendChild(body);

  const mobileBtn = document.createElement("button");
  mobileBtn.type = "button";
  mobileBtn.className = "unified-nav-mobile-btn";
  mobileBtn.textContent = "\u2630";
  mobileBtn.setAttribute("aria-label", "Open reports navigation");

  const scrim = document.createElement("div");
  scrim.className = "unified-nav-scrim";

  const closeMobile = function () {
    nav.classList.remove("is-open");
    scrim.classList.remove("is-open");
  };

  toggle.addEventListener("click", function () {
    if (isMobile) {
      closeMobile();
      return;
    }
    const collapsed = nav.classList.toggle("is-collapsed");
    document.body.classList.toggle("unified-nav-collapsed", collapsed);
    localStorage.setItem(storageKey, collapsed ? "1" : "0");
  });

  mobileBtn.addEventListener("click", function () {
    nav.classList.add("is-open");
    scrim.classList.add("is-open");
  });

  scrim.addEventListener("click", closeMobile);

  document.body.classList.add("unified-nav-enabled");
  document.body.appendChild(nav);
  document.body.appendChild(scrim);
  document.body.appendChild(mobileBtn);
})();
