import re, sys

path = r'E:\JIRA SCRIPT\report_html\team_capacity_planner.html'

with open(path, encoding='utf-8') as f:
    content = f.read()

new_section = r"""  // \u2500\u2500 Jira sync modal \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"""

new_section = (
  "  // \u2500\u2500 Jira sync modal \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n"
  "  function _jiraSyncSelectedIds() {\n"
  "    return Array.from($jiraSyncContent.querySelectorAll('input.jira-row-chk:checked'))\n"
  "      .map(cb => parseInt(cb.dataset.id, 10));\n"
  "  }\n"
  "\n"
  "  function _jiraSyncUpdatePushBtn() {\n"
  "    const selected = _jiraSyncSelectedIds();\n"
  "    if (!selected.length) {\n"
  "      $jiraSyncPush.disabled = true;\n"
  "      $jiraSyncPush.innerHTML = '<span class=\"material-symbols-outlined\" style=\"font-size:15px;vertical-align:middle\">cloud_upload</span> Push Selected to Jira';\n"
  "    } else {\n"
  "      $jiraSyncPush.disabled = false;\n"
  "      $jiraSyncPush.innerHTML = `<span class=\"material-symbols-outlined\" style=\"font-size:15px;vertical-align:middle\">cloud_upload</span> Push ${selected.length} Selected to Jira`;\n"
  "    }\n"
  "    const selCount = $jiraSyncContent.querySelector('.sel-count');\n"
  "    if (selCount) selCount.textContent = `${selected.length} selected`;\n"
  "  }\n"
  "\n"
  "  async function openJiraSyncModal() {\n"
  "    $jiraSyncModal.hidden = false;\n"
  "    $jiraSyncResult.hidden = true;\n"
  "    $jiraSyncPush.disabled = true;\n"
  "    $jiraSyncPush.innerHTML = '<span class=\"material-symbols-outlined\" style=\"font-size:15px;vertical-align:middle\">cloud_upload</span> Push Selected to Jira';\n"
  "    $jiraSyncContent.innerHTML = emptyState('hourglass_empty', 'Loading assignments\u2026');\n"
  "    try {\n"
  "      const res = await fetch('/api/team-capacity-planner/assignments');\n"
  "      const data = await res.json();\n"
  "      if (!data.ok) throw new Error(data.error || 'Failed to load');\n"
  "      const asgns = data.assignments || [];\n"
  "      if (!asgns.length) {\n"
  "        $jiraSyncContent.innerHTML = emptyState('assignment_ind', 'No assignments yet. Use the Assign button on work items.');\n"
  "        return;\n"
  "      }\n"
  "      $jiraSyncContent.innerHTML = `\n"
  "        <div class=\"jira-sel-toolbar\">\n"
  "          <input type=\"checkbox\" id=\"jira-sel-all\" title=\"Select / deselect all\" />\n"
  "          <label for=\"jira-sel-all\" style=\"cursor:pointer\">Select all</label>\n"
  "          <span class=\"sel-count\">0 selected</span>\n"
  "          <div class=\"jira-sel-actions\">\n"
  "            <button class=\"btn-sel-quick\" id=\"sel-unsynced-btn\">Select Unsynced</button>\n"
  "            <button class=\"btn-sel-quick\" id=\"sel-all-btn\">All</button>\n"
  "            <button class=\"btn-sel-quick\" id=\"sel-none-btn\">None</button>\n"
  "          </div>\n"
  "        </div>\n"
  "        <div class=\"jira-asgn-scroll\"><table class=\"jira-asgn-table\">\n"
  "          <thead><tr>\n"
  "            <th class=\"chk-col\"></th>\n"
  "            <th>Issue</th><th>Assignee</th><th>Jira Status</th><th>Updated</th>\n"
  "          </tr></thead>\n"
  "          <tbody>${asgns.map(a => {\n"
  "            const cls = a.jira_synced ? 'synced' : (a.jira_error ? 'failed' : 'unsynced');\n"
  "            const chip = a.jira_synced\n"
  "              ? '<span class=\"sync-status-chip ok\">\u2713 Synced</span>'\n"
  "              : (a.jira_error\n"
  "                ? `<span class=\"sync-status-chip error\" title=\"${esc(a.jira_error)}\">\u2717 Failed</span>`\n"
  "                : '<span class=\"sync-status-chip pending\">\u27f3 Pending</span>');\n"
  "            const errDetail = a.jira_error && !a.jira_synced ? `<br><small style=\"color:#991b1b;display:block;margin-top:2px\">${esc(a.jira_error.slice(0, 80))}</small>` : '';\n"
  "            const dt = a.updated_at_utc ? a.updated_at_utc.slice(0, 16).replace('T', ' ') : '\u2014';\n"
  "            const defaultChecked = !a.jira_synced;\n"
  "            return `<tr class=\"${cls}\" data-id=\"${a.id}\">\n"
  "              <td class=\"chk-col\"><input type=\"checkbox\" class=\"jira-row-chk\" data-id=\"${a.id}\"${defaultChecked ? ' checked' : ''} /></td>\n"
  "              <td><strong style=\"color:var(--accent)\">${esc(a.issue_key)}</strong></td>\n"
  "              <td>${esc(a.assignee_display_name || a.assignee_account_id || '\u2014')}</td>\n"
  "              <td>${chip}${errDetail}</td>\n"
  "              <td style=\"white-space:nowrap;font-size:11px;color:var(--muted)\">${esc(dt)}</td>\n"
  "            </tr>`;\n"
  "          }).join('')}</tbody>\n"
  "        </table></div>`;\n"
  "\n"
  "      _jiraSyncUpdatePushBtn();\n"
  "\n"
  "      $jiraSyncContent.querySelectorAll('input.jira-row-chk').forEach(cb => {\n"
  "        cb.addEventListener('change', _jiraSyncUpdatePushBtn);\n"
  "      });\n"
  "      const $selAll    = document.getElementById('jira-sel-all');\n"
  "      const $selAllBtn = document.getElementById('sel-all-btn');\n"
  "      const $selNone   = document.getElementById('sel-none-btn');\n"
  "      const $selUnsync = document.getElementById('sel-unsynced-btn');\n"
  "      const allCbs = () => Array.from($jiraSyncContent.querySelectorAll('input.jira-row-chk'));\n"
  "\n"
  "      $selAll.addEventListener('change', () => {\n"
  "        allCbs().forEach(cb => { cb.checked = $selAll.checked; });\n"
  "        _jiraSyncUpdatePushBtn();\n"
  "      });\n"
  "      $selAllBtn.addEventListener('click', () => {\n"
  "        allCbs().forEach(cb => { cb.checked = true; });\n"
  "        $selAll.checked = true;\n"
  "        _jiraSyncUpdatePushBtn();\n"
  "      });\n"
  "      $selNone.addEventListener('click', () => {\n"
  "        allCbs().forEach(cb => { cb.checked = false; });\n"
  "        $selAll.checked = false;\n"
  "        _jiraSyncUpdatePushBtn();\n"
  "      });\n"
  "      $selUnsync.addEventListener('click', () => {\n"
  "        allCbs().forEach(cb => {\n"
  "          const row = cb.closest('tr');\n"
  "          cb.checked = row && (row.classList.contains('unsynced') || row.classList.contains('failed'));\n"
  "        });\n"
  "        $selAll.checked = false;\n"
  "        _jiraSyncUpdatePushBtn();\n"
  "      });\n"
  "    } catch(e) {\n"
  "      $jiraSyncContent.innerHTML = emptyState('error', 'Failed to load: ' + e.message);\n"
  "    }\n"
  "  }\n"
  "\n"
  "  function closeJiraSyncModal() {\n"
  "    $jiraSyncModal.hidden = true;\n"
  "  }\n"
  "\n"
  "  async function pushAssignmentsToJira() {\n"
  "    const ids = _jiraSyncSelectedIds();\n"
  "    if (!ids.length) { showToast('No assignments selected.', 'warn'); return; }\n"
  "    $jiraSyncPush.disabled = true;\n"
  "    $jiraSyncPush.innerHTML = '<span class=\"material-symbols-outlined\" style=\"font-size:14px;vertical-align:middle;animation:spin .7s linear infinite\">sync</span> Pushing\u2026';\n"
  "    $jiraSyncResult.hidden = true;\n"
  "    try {\n"
  "      const res = await fetch('/api/team-capacity-planner/push-assignments-to-jira', {\n"
  "        method: 'POST',\n"
  "        headers: { 'Content-Type': 'application/json' },\n"
  "        body: JSON.stringify({ ids }),\n"
  "      });\n"
  "      const data = await res.json();\n"
  "      if (!data.ok) throw new Error(data.error || 'Push failed');\n"
  "      const succeeded = data.succeeded || 0;\n"
  "      const failed = (data.pushed || 0) - succeeded;\n"
  "      const msg = data.pushed === 0\n"
  "        ? 'No assignments were pushed.'\n"
  "        : succeeded > 0\n"
  "          ? `\u2713 Pushed ${succeeded} assignment(s) to Jira.${failed > 0 ? ` ${failed} failed \u2014 see table.` : ''}`\n"
  "          : `\u2717 All ${failed} push(es) failed. Check Jira connectivity.`;\n"
  "      $jiraSyncResult.hidden = false;\n"
  "      const ok = succeeded > 0 && failed === 0;\n"
  "      $jiraSyncResult.style.cssText = `margin-top:12px;padding:10px 12px;border-radius:8px;font-size:13px;background:${ok ? '#dcfce7;color:#15803d;border:1px solid #bbf7d0' : '#fff7ed;color:#92400e;border:1px solid #fcd34d'}`;\n"
  "      $jiraSyncResult.textContent = msg;\n"
  "      showToast(msg, succeeded > 0 ? 'ok' : 'warn');\n"
  "      await openJiraSyncModal();\n"
  "    } catch(e) {\n"
  "      showToast('Push failed: ' + e.message, 'error');\n"
  "      $jiraSyncPush.disabled = false;\n"
  "      $jiraSyncPush.innerHTML = '<span class=\"material-symbols-outlined\" style=\"font-size:15px;vertical-align:middle\">cloud_upload</span> Push Selected to Jira';\n"
  "    }\n"
  "  }\n"
  "\n"
)

# The start marker (the comment line itself, including trailing dashes)
start_marker = r'  // \u2500\u2500 Jira sync modal \u2500+\n'
# The end marker (event listeners comment line, kept in output)
end_marker   = r'  // \u2500\u2500 Event listeners \u2500+'

pattern = re.compile(
    r'(  // \u2500\u2500 Jira sync modal \u2500+\n).*?(?=  // \u2500\u2500 Event listeners )',
    re.DOTALL
)

if not pattern.search(content):
    print('ERROR: pattern not found in file')
    sys.exit(1)

new_content = pattern.sub(new_section, content, count=1)

with open(path, 'w', encoding='utf-8') as f:
    f.write(new_content)

# Verify
with open(path, encoding='utf-8') as f:
    check = f.read()

if '_jiraSyncSelectedIds' in check:
    print('SUCCESS: _jiraSyncSelectedIds is present in the file.')
else:
    print('FAILURE: _jiraSyncSelectedIds NOT found after replacement.')
