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
        { page_key: "introduction", title: "Introduction", href: resolveReportHref("introduction.html"), icon: "home", file: "introduction.html" },
        { page_key: "dashboard", title: "Dashboard", href: resolveReportHref("dashboard.html"), icon: "space_dashboard", file: "dashboard.html" },
        { page_key: "executive_dashboard", title: "Executive Dashboard", href: resolveReportHref("executive_dashboard.html"), icon: "analytics", file: "executive_dashboard.html" },
        { page_key: "nested_view_report", title: "Nested View Report", href: resolveReportHref("nested_view_report.html"), icon: "account_tree", file: "nested_view_report.html" },
        { page_key: "missed_entries_report", title: "Missed Entries Report", href: resolveReportHref("missed_entries.html"), icon: "event_busy", file: "missed_entries.html" },
        { page_key: "assignee_hours_report", title: "Assignee Hours Report", href: resolveReportHref("assignee_hours_report.html"), icon: "schedule", file: "assignee_hours_report.html" },
        { page_key: "employee_performance_report", title: "Employee Performance", href: resolveReportHref("employee_performance_report.html"), icon: "monitoring", file: "employee_performance_report.html" },
        { page_key: "employee_capacity_utilization", title: "Employee Capacity & Utilization", href: resolveReportHref("employee_capacity_utilization_report.html"), icon: "query_stats", file: "employee_capacity_utilization_report.html" },
        { page_key: "rlt_leave_report", title: "RLT Leave Report", href: resolveReportHref("rlt_leave_report.html"), icon: "beach_access", file: "rlt_leave_report.html" },
        { page_key: "leaves_planned_calendar", title: "Leaves Planned Calendar", href: resolveReportHref("leaves_planned_calendar.html"), icon: "calendar_month", file: "leaves_planned_calendar.html" },
        { page_key: "rnd_data_story", title: "RnD Data Story", href: resolveReportHref("rnd_data_story.html"), icon: "auto_stories", file: "rnd_data_story.html" },
        { page_key: "planned_rmis_report", title: "Planned RMIs", href: resolveReportHref("planned_rmis_report.html"), icon: "assignment_turned_in", file: "planned_rmis_report.html" },
        { page_key: "phase_rmi_gantt_report", title: "Phase RMI Gantt", href: resolveReportHref("phase_rmi_gantt_report.html"), icon: "view_timeline", file: "phase_rmi_gantt_report.html" },
        { page_key: "rmi_jira_gantt_report", title: "RMI Jira Gantt", href: resolveReportHref("rmi_jira_gantt_report.html"), icon: "view_timeline", file: "rmi_jira_gantt_report.html" },
        { page_key: "approved_vs_planned_hours_report", title: "Approved vs Planned Hours Report", href: resolveReportHref("approved_vs_planned_hours_report.html"), icon: "analytics", file: "approved_vs_planned_hours_report.html" },
        { page_key: "planned_actual_table_view", title: "Planned vs Actual Table View", href: resolveReportHref("planned_actual_table_view.html"), icon: "table_view", file: "planned_actual_table_view.html" },
        { page_key: "original_estimates_hierarchy_report", title: "Epic Estimate Report", href: resolveReportHref("original_estimates_hierarchy_report.html"), icon: "schema", file: "original_estimates_hierarchy_report.html" },
        { page_key: "delayed_epic_chain_gantt", title: "Delayed Epic Chain Gantt", href: resolveReportHref("delayed_epic_chain_gantt_report.html"), icon: "timeline", file: "delayed_epic_chain_gantt_report.html" },
        { page_key: "monthly_epic_plan_progress", title: "Monthly Epic Plan vs Actual", href: resolveReportHref("monthly_epic_plan_progress_report.html"), icon: "event_available", file: "monthly_epic_plan_progress_report.html" },
        { page_key: "support_center_report", title: "Support Center", href: resolveReportHref("support_center_report.html"), icon: "support_agent", file: "support_center_report.html" },

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
        { page_key: "executive_dashboard_settings", title: "Executive Dashboard Settings", href: "/settings/executive-dashboard", icon: "tune", path: "/settings/executive-dashboard" },
        { page_key: "report_entities", title: "Report Entities", href: "/settings/report-entities", icon: "dataset", path: "/settings/report-entities" },
        { page_key: "manage_fields", title: "Manage Fields", href: "/settings/manage-fields", icon: "list_alt", path: "/settings/manage-fields" },
        { page_key: "projects", title: "Projects", href: "/settings/projects", icon: "work", path: "/settings/projects" },
        { page_key: "epic_dropdowns", title: "Epic Dropdowns", href: "/settings/epics-dropdown-options", icon: "arrow_drop_down_circle", path: "/settings/epics-dropdown-options" },
        { page_key: "epic_phases", title: "Epic Phases", href: "/settings/epic-phases", icon: "alt_route", path: "/settings/epic-phases" },
        { page_key: "epics_planner", title: "Epics Planner", href: "/settings/epics-management", icon: "event_note", path: "/settings/epics-management" },
        { page_key: "epics_planner_import", title: "Epics Planner Import", href: "/settings/epics-management/import", icon: "upload_file", path: "/settings/epics-management/import" },
        { page_key: "ipp_meeting_planner", title: "IPP Meeting Planner", href: "/settings/ipp-meeting-planner", icon: "groups", path: "/settings/ipp-meeting-planner" },
        { page_key: "rnd_muscle_utilization", title: "RnD Muscle Utilization", href: "/settings/rnd-muscle-utilization", icon: "hub", path: "/settings/rnd-muscle-utilization" },
        { page_key: "page_categories", title: "Page Categories", href: "/settings/page-categories", icon: "category", path: "/settings/page-categories" },
        { page_key: "canonical_refresh_settings", title: "Colossal Refresh", href: "/settings/canonical-refresh", icon: "sync", path: "/settings/canonical-refresh" },
        { page_key: "team_capacity_planner", title: "Team Capacity Planner", href: "/settings/team-capacity-planner", icon: "group_work", path: "/settings/team-capacity-planner" },
        { page_key: "sql_console", title: "SQL Console", href: "/settings/sql-console", icon: "query_stats", path: "/settings/sql-console" },
        { page_key: "product_releases", title: "Product Releases", href: "/settings/product-releases", icon: "rocket_launch", path: "/settings/product-releases" },
        { page_key: "seating_planner", title: "Seating Planner", href: "/settings/seating-planner", icon: "chair", path: "/settings/seating-planner" }
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
    let NAV_SECTIONS = await resolveNavSections();
    let eprIdentity = null;
    if (window.location.protocol !== 'file:') {
      try {
        const authResponse = await fetch('/api/auth/me', {cache:'no-store'});
        const auth = await authResponse.json();
        if (auth.enabled && !auth.authenticated) { window.location.replace('/login?next='+encodeURIComponent(location.pathname+location.search)); return; }
        if (auth.enabled) {
          eprIdentity = auth.user;
          if (!eprIdentity.active || !eprIdentity.approved) {location.replace('/access-pending');return;}
          if (eprIdentity.role !== 'admin') NAV_SECTIONS = NAV_SECTIONS.filter(section=>section.id!=='admin-settings');
          else {
            let admin = NAV_SECTIONS.find(section=>section.id==='admin-settings');
            if (!admin) {admin={id:'admin-settings',mainLabel:'Admin Settings',mainIcon:'admin_panel_settings',title:'Admin Settings',items:[],categories:[]};NAV_SECTIONS.push(admin);}
            admin.items.push({page_key:'users_access',title:'Users & access',href:'/settings/users',path:'/settings/users',icon:'manage_accounts'});
          }
        }
      } catch (_error) { /* File previews and older local servers retain their navigation. Server authorization remains authoritative. */ }
    }
    const adminNav = NAV_SECTIONS.find((section) => section.id === "admin-settings");
    if (adminNav && !(adminNav.items || []).some((item) => item.page_key === "usage_analytics")) {
      adminNav.items.push({ page_key: "usage_analytics", title: "Usage Analytics", href: "/settings/usage-analytics", path: "/settings/usage-analytics", icon: "insights" });
    }
    const MY_DAY_ITEM = { page_key: "my_day", title: "My Day", href: resolveReportHref("my_day.html"), icon: "wb_sunny", file: "my_day.html" };
    const storageKey = "unified-report-nav-collapsed";
    const groupsStorageKey = "unified-report-nav-groups";
    const mobileQuery = window.matchMedia("(max-width: 959px)");
    const isMobile = function () { return mobileQuery.matches; };
    const currentPath = String(window.location.pathname || "").replace(/\\/g, "/").toLowerCase();
    const currentFile = (window.location.pathname.split("/").pop() || "").toLowerCase();
    ensureMaterialSymbolsLoaded();

    function readStore(key) {
      try { return localStorage.getItem(key); } catch (_err) { return null; }
    }
    function writeStore(key, value) {
      try { localStorage.setItem(key, value); } catch (_err) { /* storage can be blocked; nav still works */ }
    }

    function symbol(name, className) {
      const el = document.createElement("span");
      el.className = (className ? className + " " : "") + "material-symbols-outlined";
      el.setAttribute("aria-hidden", "true");
      el.textContent = name;
      return el;
    }

    const nav = document.createElement("aside");
    nav.className = "unified-nav";
    nav.id = "unified-nav";
    nav.setAttribute("aria-label", "Reports navigation");

    if (!isMobile() && readStore(storageKey) === "1") {
      nav.classList.add("is-collapsed");
      document.body.classList.add("unified-nav-collapsed");
    }

    const head = document.createElement("div");
    head.className = "unified-nav-head";

    const brand = document.createElement("div");
    brand.className = "unified-nav-brand";
    brand.innerHTML = '<h2 class="unified-nav-title">EPR Tool</h2><div class="unified-nav-sub">Employee Performance Reporting Tool - Octopus Digital</div>';

    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "unified-nav-toggle";
    toggle.setAttribute("aria-controls", "unified-nav");
    const toggleIcon = symbol("left_panel_close");
    toggle.appendChild(toggleIcon);

    head.appendChild(brand);
    head.appendChild(toggle);
    nav.appendChild(head);

    const searchWrap = document.createElement("label");
    searchWrap.className = "unified-nav-search";
    searchWrap.appendChild(symbol("search", "unified-nav-search-icon"));
    const searchInput = document.createElement("input");
    searchInput.type = "search";
    searchInput.placeholder = "Find a page";
    searchInput.setAttribute("aria-label", "Find a page in navigation");
    searchInput.autocomplete = "off";
    searchWrap.appendChild(searchInput);
    nav.appendChild(searchWrap);

    const body = document.createElement("div");
    body.className = "unified-nav-body";

    const mainNav = document.createElement("nav");
    mainNav.className = "unified-nav-main";
    mainNav.setAttribute("aria-label", "Main navigation");

    const emptyState = document.createElement("p");
    emptyState.className = "unified-nav-empty";
    emptyState.textContent = "No pages match.";
    emptyState.hidden = true;

    const reportsSection = NAV_SECTIONS.find((section) => section.id === "reports") || NAV_SECTIONS[0];
    const adminSection = NAV_SECTIONS.find((section) => section.id === "admin-settings");

    let savedGroups = null;
    try { savedGroups = JSON.parse(readStore(groupsStorageKey) || "null"); } catch (_err) { savedGroups = null; }
    const expandedGroups = new Set(Array.isArray(savedGroups) ? savedGroups : []);
    const groupEls = [];

    function persistGroups() {
      writeStore(groupsStorageKey, JSON.stringify(Array.from(expandedGroups)));
    }

    function isItemActive(item) {
      const isActiveFile = item.file && currentFile === String(item.file || "").toLowerCase();
      const isActivePath = item.path && currentPath.indexOf(String(item.path || "").toLowerCase()) === 0;
      return !!(isActiveFile || isActivePath);
    }

    function createLink(item) {
      const link = document.createElement("a");
      link.className = "unified-nav-main-link";
      link.href = item.href;
      link.title = item.title;
      link.dataset.navTitle = String(item.title || "").toLowerCase();
      if (isItemActive(item)) {
        link.classList.add("is-active");
        link.setAttribute("aria-current", "page");
      }
      const label = document.createElement("span");
      label.className = "unified-nav-main-label";
      label.textContent = item.title;
      link.appendChild(symbol(item.icon || "description", "unified-nav-main-icon"));
      link.appendChild(label);
      return link;
    }

    function categoryKey(prefix, group, index) {
      const id = Number(group && group.id ? group.id : 0);
      if (id > 0) return prefix + ":" + String(id);
      return prefix + ":" + String(group && group.name ? group.name : index);
    }

    let groupSeq = 0;
    // A group is a disclosure: header button + animated item panel. Groups nest (admin categories sit inside Admin Settings).
    function createGroup(key, labelText, iconName, items, children, options) {
      const opts = options || {};
      const groupEl = document.createElement("div");
      groupEl.className = "unified-nav-group" + (opts.nested ? " is-nested" : "") + (opts.prominent ? " is-prominent" : "");
      groupEl.dataset.groupKey = key;

      const panelId = "unified-nav-group-" + String(++groupSeq);
      const button = document.createElement("button");
      button.type = "button";
      button.className = "unified-nav-group-toggle";
      button.title = labelText;
      button.setAttribute("aria-controls", panelId);

      const label = document.createElement("span");
      label.className = "unified-nav-group-label";
      label.textContent = labelText;

      const total = items.length + children.reduce((sum, child) => sum + child.count, 0);
      const count = document.createElement("span");
      count.className = "unified-nav-group-count";
      count.textContent = String(total);

      button.appendChild(symbol(iconName || "folder", "unified-nav-group-icon"));
      button.appendChild(label);
      button.appendChild(count);
      button.appendChild(symbol("expand_more", "unified-nav-group-chevron"));

      const panel = document.createElement("div");
      panel.className = "unified-nav-group-panel";
      panel.id = panelId;
      const inner = document.createElement("div");
      inner.className = "unified-nav-group-items";
      items.forEach((item) => inner.appendChild(createLink(item)));
      children.forEach((child) => inner.appendChild(child.el));
      panel.appendChild(inner);

      groupEl.appendChild(button);
      groupEl.appendChild(panel);

      if (inner.querySelector(".is-active")) {
        groupEl.classList.add("has-active");
        expandedGroups.add(key);
      }

      const record = { key: key, el: groupEl, button: button, count: total };
      groupEls.push(record);

      button.addEventListener("click", function () {
        if (nav.classList.contains("is-collapsed")) {
          setCollapsed(false);
          expandedGroups.add(key);
        } else if (expandedGroups.has(key)) {
          expandedGroups.delete(key);
        } else {
          expandedGroups.add(key);
        }
        persistGroups();
        syncGroups();
        if (expandedGroups.has(key)) revealAfterExpand(groupEl);
      });
      return record;
    }

    // Groups near the bottom (Admin Settings) would otherwise open below the fold with no visible change.
    function revealAfterExpand(groupEl) {
      window.setTimeout(function () {
        const box = mainNav.getBoundingClientRect();
        const rect = groupEl.getBoundingClientRect();
        const overflow = rect.bottom - box.bottom;
        if (overflow > 0) mainNav.scrollBy({ top: Math.min(overflow + 8, rect.top - box.top), behavior: "smooth" });
      }, 240);
    }

    function syncGroups() {
      const filtering = !!searchInput.value.trim();
      groupEls.forEach((group) => {
        const open = filtering ? !group.el.classList.contains("is-filtered-out") : expandedGroups.has(group.key);
        group.el.classList.toggle("is-open", open);
        group.button.setAttribute("aria-expanded", open ? "true" : "false");
      });
    }

    function buildCategoryGroups(prefix, categories, nested) {
      return (Array.isArray(categories) ? categories : [])
        .map((group, idx) => {
          const name = String(group.name || "").trim();
          const items = Array.isArray(group.items) ? group.items : [];
          if (!name || !items.length) return null;
          return createGroup(categoryKey(prefix, group, idx), name, String(group.icon_name || "folder"), items, [], { nested: nested });
        })
        .filter(Boolean);
    }

    function renderMainNav() {
      mainNav.innerHTML = "";
      groupEls.length = 0;
      const myDayLink = createLink(MY_DAY_ITEM);
      myDayLink.classList.add("is-pinned");
      mainNav.appendChild(myDayLink);
      (reportsSection.items || []).filter((item) => item.page_key !== "my_day").forEach((item) => mainNav.appendChild(createLink(item)));
      const reportGroups = buildCategoryGroups("reports", reportsSection.categories, false);
      if (!savedGroups && reportGroups.length && !reportGroups.some((g) => expandedGroups.has(g.key))) {
        expandedGroups.add(reportGroups[0].key);
      }
      reportGroups.forEach((group) => mainNav.appendChild(group.el));

      if (adminSection) {
        const divider = document.createElement("div");
        divider.className = "unified-nav-divider";
        mainNav.appendChild(divider);
        const adminChildren = buildCategoryGroups("admin", adminSection.categories, true);
        const adminGroup = createGroup("admin-settings", adminSection.mainLabel, adminSection.mainIcon,
          adminSection.items || [], adminChildren, { prominent: true });
        mainNav.appendChild(adminGroup.el);
      }
      mainNav.appendChild(emptyState);
      syncGroups();
    }

    function applyFilter() {
      const query = searchInput.value.trim().toLowerCase();
      let visible = 0;
      mainNav.querySelectorAll(".unified-nav-main-link").forEach((link) => {
        const match = !query || link.dataset.navTitle.indexOf(query) !== -1;
        link.hidden = !match;
        if (match) visible += 1;
      });
      groupEls.forEach((group) => {
        const hasMatch = !query || !!group.el.querySelector(".unified-nav-main-link:not([hidden])");
        group.el.classList.toggle("is-filtered-out", !hasMatch);
      });
      emptyState.hidden = visible > 0;
      nav.classList.toggle("is-filtering", !!query);
      syncGroups();
    }

    searchInput.addEventListener("input", applyFilter);
    searchInput.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && searchInput.value) {
        event.stopPropagation();
        searchInput.value = "";
        applyFilter();
      } else if (event.key === "Enter") {
        const first = mainNav.querySelector(".unified-nav-main-link:not([hidden])");
        if (first) window.location.href = first.href;
      }
    });

    renderMainNav();
    body.appendChild(mainNav);
    nav.appendChild(body);

    if (eprIdentity) {
      const account = document.createElement("div");
      account.className = "epr-nav-account";
      const avatar = document.createElement("span");
      avatar.className = "epr-nav-avatar";
      avatar.textContent = (eprIdentity.display_name || "E")[0].toUpperCase();
      if (/^https:\/\//i.test(eprIdentity.avatar_url || "")) {
        const image = document.createElement("img");
        image.src = eprIdentity.avatar_url;
        image.alt = "";
        avatar.textContent = "";
        avatar.appendChild(image);
      }
      const name = document.createElement("strong");
      name.textContent = eprIdentity.display_name || "Atlassian user";
      const role = document.createElement("small");
      role.textContent = eprIdentity.role === "admin" ? "Administrator" : "Member";
      const identity = document.createElement("div");
      identity.className = "epr-nav-identity";
      identity.title = name.textContent;
      identity.append(name, role);

      const actions = document.createElement("div");
      actions.className = "epr-nav-actions";
      const mine = document.createElement("a");
      mine.className = "epr-nav-action";
      mine.href = "/my-availability";
      mine.title = "My availability";
      const mineLabel = document.createElement("span");
      mineLabel.textContent = "My availability";
      mine.append(symbol("event_available"), mineLabel);

      const logout = document.createElement("button");
      logout.type = "button";
      logout.className = "epr-nav-action epr-nav-signout";
      logout.title = "Sign out";
      const logoutLabel = document.createElement("span");
      logoutLabel.textContent = "Sign out";
      logout.append(symbol("logout"), logoutLabel);
      logout.addEventListener("click", async function () {
        logout.disabled = true;
        logoutLabel.textContent = "Signing out…";
        try {
          const response = await fetch("/api/auth/logout", { method: "POST", headers: { "Content-Type": "application/json", "X-EPR-CSRF": eprIdentity.csrf_token }, body: "{}" });
          if (response.ok) { location.replace("/login"); return; }
        } catch (_err) { /* falls through to the retry state below */ }
        logout.disabled = false;
        logoutLabel.textContent = "Sign out failed — retry";
      });

      actions.append(mine, logout);
      account.append(avatar, identity, actions);
      nav.appendChild(account);
    }

    const mobileBtn = document.createElement("button");
    mobileBtn.type = "button";
    mobileBtn.className = "unified-nav-mobile-btn";
    mobileBtn.setAttribute("aria-label", "Open navigation");
    mobileBtn.setAttribute("aria-controls", "unified-nav");
    mobileBtn.setAttribute("aria-expanded", "false");
    mobileBtn.appendChild(symbol("menu"));

    const scrim = document.createElement("div");
    scrim.className = "unified-nav-scrim";

    function syncToggleLabel() {
      if (isMobile()) {
        toggleIcon.textContent = "close";
        toggle.setAttribute("aria-label", "Close navigation");
        toggle.removeAttribute("aria-expanded");
        return;
      }
      const collapsed = nav.classList.contains("is-collapsed");
      toggleIcon.textContent = collapsed ? "left_panel_open" : "left_panel_close";
      toggle.setAttribute("aria-label", collapsed ? "Expand navigation" : "Collapse navigation");
      toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }

    function setCollapsed(collapsed) {
      nav.classList.toggle("is-collapsed", collapsed);
      document.body.classList.toggle("unified-nav-collapsed", collapsed);
      writeStore(storageKey, collapsed ? "1" : "0");
      syncToggleLabel();
    }

    function openMobile() {
      nav.classList.add("is-open");
      scrim.classList.add("is-open");
      mobileBtn.setAttribute("aria-expanded", "true");
      const active = mainNav.querySelector(".is-active");
      if (active) active.scrollIntoView({ block: "center" });
      toggle.focus();
    }

    function closeMobile() {
      if (!nav.classList.contains("is-open")) return;
      nav.classList.remove("is-open");
      scrim.classList.remove("is-open");
      mobileBtn.setAttribute("aria-expanded", "false");
      mobileBtn.focus();
    }

    toggle.addEventListener("click", function () {
      if (isMobile()) closeMobile();
      else setCollapsed(!nav.classList.contains("is-collapsed"));
    });
    mobileBtn.addEventListener("click", openMobile);
    scrim.addEventListener("click", closeMobile);
    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && isMobile()) closeMobile();
    });
    mobileQuery.addEventListener("change", function () {
      if (isMobile()) {
        nav.classList.remove("is-collapsed");
        document.body.classList.remove("unified-nav-collapsed");
      } else {
        closeMobile();
        if (readStore(storageKey) === "1") {
          nav.classList.add("is-collapsed");
          document.body.classList.add("unified-nav-collapsed");
        }
      }
      syncToggleLabel();
    });
    syncToggleLabel();

    document.body.classList.add("unified-nav-enabled");
    document.body.appendChild(nav);
    document.body.appendChild(scrim);
    document.body.appendChild(mobileBtn);

    const activeLink = mainNav.querySelector(".is-active");
    if (activeLink && !isMobile()) activeLink.scrollIntoView({ block: "nearest" });
  }

  initUnifiedNav();
})();
