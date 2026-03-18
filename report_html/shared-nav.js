(function () {
  function ensureMaterialSymbolsLoaded() {
    if (document.getElementById("unified-nav-material-symbols")) return;
    const link = document.createElement("link");
    link.id = "unified-nav-material-symbols";
    link.rel = "stylesheet";
    link.href = window.location.protocol === "file:" ? "./material-symbols.css" : "/material-symbols.css";
    document.head.appendChild(link);
  }

  function resolveReportHref(fileName) {
    if (window.location.protocol === "file:") return "report_html/" + fileName;
    return "/" + fileName;
  }

  const STATIC_NAV_SECTIONS = [
    {
      id: "reports",
      mainLabel: "Reports",
      mainIcon: "description",
      title: "Reports",
      items: [
        { page_key: "dashboard", title: "Dashboard", href: resolveReportHref("dashboard.html"), icon: "space_dashboard", file: "dashboard.html" },
        { page_key: "nested_view_report", title: "Nested View Report", href: resolveReportHref("nested_view_report.html"), icon: "account_tree", file: "nested_view_report.html" },
        { page_key: "missed_entries_report", title: "Missed Entries Report", href: resolveReportHref("missed_entries.html"), icon: "event_busy", file: "missed_entries.html" },
        { page_key: "assignee_hours_report", title: "Assignee Hours Report", href: resolveReportHref("assignee_hours_report.html"), icon: "schedule", file: "assignee_hours_report.html" },
        { page_key: "employee_performance_report", title: "Employee Performance", href: resolveReportHref("employee_performance_report.html"), icon: "monitoring", file: "employee_performance_report.html" },
        { page_key: "rlt_leave_report", title: "RLT Leave Report", href: resolveReportHref("rlt_leave_report.html"), icon: "beach_access", file: "rlt_leave_report.html" },
        { page_key: "leaves_planned_calendar", title: "Leaves Planned Calendar", href: resolveReportHref("leaves_planned_calendar.html"), icon: "calendar_month", file: "leaves_planned_calendar.html" },
        { page_key: "rnd_data_story", title: "RnD Data Story", href: resolveReportHref("rnd_data_story.html"), icon: "auto_stories", file: "rnd_data_story.html" },
        { page_key: "planned_rmis_report", title: "Planned RMIs", href: resolveReportHref("planned_rmis_report.html"), icon: "assignment_turned_in", file: "planned_rmis_report.html" },
        { page_key: "phase_rmi_gantt_report", title: "Phase RMI Gantt", href: resolveReportHref("phase_rmi_gantt_report.html"), icon: "view_timeline", file: "phase_rmi_gantt_report.html" },
        { page_key: "approved_vs_planned_hours_report", title: "Approved vs Planned Hours Report", href: resolveReportHref("approved_vs_planned_hours_report.html"), icon: "analytics", file: "approved_vs_planned_hours_report.html" },
        { page_key: "planned_actual_table_view", title: "Planned vs Actual Table View", href: resolveReportHref("planned_actual_table_view.html"), icon: "table_view", file: "planned_actual_table_view.html" },
        { page_key: "original_estimates_hierarchy_report", title: "Epic Estimate Report", href: resolveReportHref("original_estimates_hierarchy_report.html"), icon: "schema", file: "original_estimates_hierarchy_report.html" },
        { page_key: "ipp_meeting_dashboard", title: "IPP Meeting Dashboard", href: resolveReportHref("ipp_meeting_dashboard.html"), icon: "groups", file: "ipp_meeting_dashboard.html" }
      ],
      categories: []
    },
    {
      id: "admin-settings",
      mainLabel: "Admin Settings",
      mainIcon: "admin_panel_settings",
      title: "Admin Settings",
      items: [
        { page_key: "capacity_settings", title: "Capacity Settings", href: "/settings/capacity", icon: "tune", path: "/settings/capacity" },
        { page_key: "performance_settings", title: "Performance Settings", href: "/settings/performance", icon: "speed", path: "/settings/performance" },
        { page_key: "report_entities", title: "Report Entities", href: "/settings/report-entities", icon: "dataset", path: "/settings/report-entities" },
        { page_key: "manage_fields", title: "Manage Fields", href: "/settings/manage-fields", icon: "list_alt", path: "/settings/manage-fields" },
        { page_key: "projects", title: "Projects", href: "/settings/projects", icon: "work", path: "/settings/projects" },
        { page_key: "epic_dropdowns", title: "Epic Dropdowns", href: "/settings/epics-dropdown-options", icon: "arrow_drop_down_circle", path: "/settings/epics-dropdown-options" },
        { page_key: "epic_phases", title: "Epic Phases", href: "/settings/epic-phases", icon: "alt_route", path: "/settings/epic-phases" },
        { page_key: "epics_planner", title: "Epics Planner", href: "/settings/epics-management", icon: "event_note", path: "/settings/epics-management" },
        { page_key: "page_categories", title: "Page Categories", href: "/settings/page-categories", icon: "category", path: "/settings/page-categories" },
        { page_key: "sql_console", title: "SQL Console", href: "/settings/sql-console", icon: "query_stats", path: "/settings/sql-console" }
      ],
      categories: []
    }
  ];

  function cloneStaticNavSections() {
    return JSON.parse(JSON.stringify(STATIC_NAV_SECTIONS));
  }

  function applyCatalogTitles(sections, pageCatalog) {
    const catalog = Array.isArray(pageCatalog) ? pageCatalog : [];
    const titleByPageKey = new Map();
    catalog.forEach((item) => {
      const pageKey = String(item.page_key || "");
      if (!pageKey) return;
      const title = String(item.title || "").trim();
      if (!title) return;
      titleByPageKey.set(pageKey, title);
    });
    sections.forEach((section) => {
      (Array.isArray(section.items) ? section.items : []).forEach((item) => {
        const title = titleByPageKey.get(String(item.page_key || ""));
        if (title) item.title = title;
      });
    });
    return sections;
  }

  function normalizeNavItem(rawItem) {
    const item = rawItem || {};
    const out = {
      page_key: String(item.page_key || ""),
      title: String(item.title || ""),
      href: String(item.href || ""),
      icon: String(item.icon || "description")
    };
    if (item.file) out.file = String(item.file);
    if (item.path) out.path = String(item.path);
    return out;
  }

  async function resolveNavSections() {
    const staticSections = cloneStaticNavSections();
    if (window.location.protocol === "file:") return staticSections;
    try {
      const response = await fetch("/api/page-categories", { cache: "no-store" });
      if (!response.ok) return staticSections;
      const body = await response.json().catch(() => ({}));
      applyCatalogTitles(staticSections, body && body.page_catalog);
      const nav = body && body.navigation ? body.navigation : {};
      if (!nav || !nav.enabled) return staticSections;

      const reportsSection = staticSections.find((section) => section.id === "reports") || staticSections[0];
      const adminSection = staticSections.find((section) => section.id === "admin-settings") || staticSections[1];

      const reportsRaw = nav.reports || {};
      const reportsItems = Array.isArray(reportsRaw.items) ? reportsRaw.items.map(normalizeNavItem) : [];
      const reportsCategories = Array.isArray(reportsRaw.categories)
        ? reportsRaw.categories.map((group) => ({
            id: Number(group.id || 0),
            name: String(group.name || ""),
            icon_name: String(group.icon_name || "folder"),
            items: Array.isArray(group.items) ? group.items.map(normalizeNavItem) : []
          }))
        : [];

      reportsSection.items = reportsItems;
      reportsSection.categories = reportsCategories;

      const adminRaw = nav.admin_settings || {};
      adminSection.items = Array.isArray(adminRaw.items) ? adminRaw.items.map(normalizeNavItem) : [];
      adminSection.categories = Array.isArray(adminRaw.categories)
        ? adminRaw.categories.map((group) => ({
            id: Number(group.id || 0),
            name: String(group.name || ""),
            icon_name: String(group.icon_name || "folder"),
            items: Array.isArray(group.items) ? group.items.map(normalizeNavItem) : []
          }))
        : [];

      return staticSections;
    } catch (_err) {
      return staticSections;
    }
  }

  async function initUnifiedNav() {
    const NAV_SECTIONS = await resolveNavSections();
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

    const reportsSection = NAV_SECTIONS.find((section) => section.id === "reports") || NAV_SECTIONS[0];
    const adminSection = NAV_SECTIONS.find((section) => section.id === "admin-settings");
    let isAdminExpanded = false;
    const mainButtonsBySectionId = {};
    const expandedReportCategoryKeys = new Set();
    const expandedAdminCategoryKeys = new Set();

    function isItemActive(item) {
      const isActiveFile = item.file && currentFile === String(item.file || "").toLowerCase();
      const isActivePath = item.path && currentPath.indexOf(String(item.path || "").toLowerCase()) === 0;
      return !!(isActiveFile || isActivePath);
    }

    function syncSecondaryNavState() {
      const isSecondaryOpen = Boolean(adminSection && isAdminExpanded);
      nav.classList.toggle("has-secondary", isSecondaryOpen);
      document.body.classList.toggle("unified-nav-secondary-open", isSecondaryOpen);
    }

    function createMainLink(item) {
      const link = document.createElement("a");
      link.className = "unified-nav-main-link";
      link.href = item.href;
      if (isItemActive(item)) {
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
      return link;
    }

    function createMainGroupTitle(text) {
      const el = document.createElement("div");
      el.className = "unified-nav-section-title";
      el.style.margin = "10px 8px 6px";
      el.style.fontSize = "0.75rem";
      el.style.textTransform = "uppercase";
      el.style.letterSpacing = ".04em";
      el.textContent = text;
      return el;
    }

    function categoryKey(prefix, group, index) {
      const id = Number(group && group.id ? group.id : 0);
      if (id > 0) return prefix + ":" + String(id);
      return prefix + ":" + String(group && group.name ? group.name : index);
    }

    function ensureDefaultExpanded(groups, expandedSet, prefix) {
      if (!Array.isArray(groups) || groups.length === 0 || expandedSet.size > 0) return;
      expandedSet.add(categoryKey(prefix, groups[0], 0));
    }

    function createCategoryToggleButton(labelText, iconName, expanded, onToggle) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "unified-nav-main-link";
      button.style.margin = "10px 0 6px";
      button.style.justifyContent = "space-between";
      button.style.gap = "8px";

      const labelWrap = document.createElement("span");
      labelWrap.style.display = "inline-flex";
      labelWrap.style.alignItems = "center";
      labelWrap.style.gap = "8px";

      const icon = document.createElement("span");
      icon.className = "material-symbols-outlined";
      icon.textContent = String(iconName || "folder");
      icon.style.fontSize = "1rem";
      icon.style.color = "inherit";
      icon.style.opacity = "0.95";

      const label = document.createElement("span");
      label.className = "unified-nav-main-label";
      label.textContent = labelText;
      label.style.fontSize = "0.78rem";
      label.style.textTransform = "uppercase";
      label.style.letterSpacing = ".04em";

      const chevron = document.createElement("span");
      chevron.className = "material-symbols-outlined";
      chevron.textContent = expanded ? "expand_more" : "chevron_right";
      chevron.style.fontSize = "1rem";
      chevron.style.color = "inherit";
      chevron.style.opacity = "0.8";

      labelWrap.appendChild(icon);
      labelWrap.appendChild(label);
      button.appendChild(labelWrap);
      button.appendChild(chevron);
      button.addEventListener("click", onToggle);
      return button;
    }

    function renderMainNav() {
      mainNav.innerHTML = "";
      (reportsSection.items || []).forEach((item) => {
        mainNav.appendChild(createMainLink(item));
      });
      const reportCategories = Array.isArray(reportsSection.categories) ? reportsSection.categories : [];
      ensureDefaultExpanded(reportCategories, expandedReportCategoryKeys, "reports");
      reportCategories.forEach((group, idx) => {
        const title = String(group.name || "").trim();
        if (!title) return;
        const key = categoryKey("reports", group, idx);
        const isExpanded = expandedReportCategoryKeys.has(key);
        const toggleBtn = createCategoryToggleButton(title, String(group.icon_name || "folder"), isExpanded, function () {
          if (expandedReportCategoryKeys.has(key)) expandedReportCategoryKeys.delete(key);
          else expandedReportCategoryKeys.add(key);
          renderMainNav();
        });
        mainNav.appendChild(toggleBtn);

        const categoryItemsWrap = document.createElement("div");
        categoryItemsWrap.style.display = isExpanded ? "block" : "none";
        (group.items || []).forEach((item) => {
          categoryItemsWrap.appendChild(createMainLink(item));
        });
        mainNav.appendChild(categoryItemsWrap);
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

    function createSecondaryLink(item) {
      const link = document.createElement("a");
      link.className = "unified-nav-link";
      link.href = item.href;
      if (isItemActive(item)) {
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
      return link;
    }

    function createSecondaryCategorySection(name, iconName, items, isExpanded, onToggle) {
      const sectionEl = document.createElement("section");
      sectionEl.className = "unified-nav-section";
      const sectionTitleBtn = createCategoryToggleButton(name, iconName, isExpanded, onToggle);
      sectionTitleBtn.style.margin = "0 0 6px";
      sectionEl.appendChild(sectionTitleBtn);
      const itemsWrap = document.createElement("div");
      itemsWrap.style.display = isExpanded ? "block" : "none";
      items.forEach((item) => itemsWrap.appendChild(createSecondaryLink(item)));
      sectionEl.appendChild(itemsWrap);
      return sectionEl;
    }

    function renderSecondaryNav() {
      list.innerHTML = "";
      const adminBtn = adminSection ? mainButtonsBySectionId[adminSection.id] : null;
      if (adminBtn) adminBtn.classList.toggle("is-active", isAdminExpanded);
      if (!adminSection || !isAdminExpanded) {
        list.classList.add("is-hidden");
        syncSecondaryNavState();
        return;
      }
      list.classList.remove("is-hidden");

      const defaultSection = document.createElement("section");
      defaultSection.className = "unified-nav-section";

      const defaultTitle = document.createElement("h3");
      defaultTitle.className = "unified-nav-section-title";
      defaultTitle.textContent = adminSection.title;
      defaultSection.appendChild(defaultTitle);

      (adminSection.items || []).forEach((item) => defaultSection.appendChild(createSecondaryLink(item)));
      list.appendChild(defaultSection);

      const adminCategories = Array.isArray(adminSection.categories) ? adminSection.categories : [];
      ensureDefaultExpanded(adminCategories, expandedAdminCategoryKeys, "admin");
      adminCategories.forEach((group, idx) => {
        const name = String(group.name || "").trim();
        if (!name) return;
        const items = Array.isArray(group.items) ? group.items : [];
        if (!items.length) return;
        const key = categoryKey("admin", group, idx);
        const isExpanded = expandedAdminCategoryKeys.has(key);
        list.appendChild(
          createSecondaryCategorySection(
            name,
            String(group.icon_name || "folder"),
            items,
            isExpanded,
            function () {
              if (expandedAdminCategoryKeys.has(key)) expandedAdminCategoryKeys.delete(key);
              else expandedAdminCategoryKeys.add(key);
              renderSecondaryNav();
            }
          )
        );
      });

      syncSecondaryNavState();
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
    syncSecondaryNavState();
    document.body.appendChild(nav);
    document.body.appendChild(scrim);
    document.body.appendChild(mobileBtn);
  }

  initUnifiedNav();
})();
