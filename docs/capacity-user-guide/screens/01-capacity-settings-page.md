# Capacity Settings Page

## Screen

- Name: Capacity Profile Settings
- Route: `/settings/capacity`
- Purpose: Create, update, load, and delete saved capacity profiles by date range.

## Sections

### Profile Editor

#### Area: Saved Profile Selection

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Saved Profiles (`profile-select`) | Dropdown | No | `No saved profiles found` when empty | Lists saved profiles as `from_date to to_date`. | Disabled when no profiles exist. |

#### Area: Profile Inputs

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| From Date (`from-date`) | Date | Yes | Empty | Start date of profile range. | Must be valid ISO date; `to_date >= from_date`. |
| To Date (`to-date`) | Date | Yes | Empty | End date of profile range. | Must be valid ISO date; `to_date >= from_date`. |
| Employees (`employees`) | Number | Yes | `0` | Team size used for capacity calculation. | Integer; must be `>= 0`. |
| Standard Hours/Day (`std-hours`) | Number | Yes | `8` | Hours per non-Ramadan working day. | Must be `> 0`. |
| Ramadan Hours/Day (`ramadan-hours`) | Number | Yes | `6.5` | Hours per Ramadan working day. | Must be `> 0`. |
| Ramadan Start (`ramadan-start`) | Date | Conditional | Empty | Start of Ramadan period. | Must be set with Ramadan End or both empty. |
| Ramadan End (`ramadan-end`) | Date | Conditional | Empty | End of Ramadan period. | Must be set with Ramadan Start; must be `>= Ramadan Start`. |
| Holiday Date Picker (`holiday-date-picker`) | Date | No | Empty | Calendar input used to choose one holiday date at a time. | Date is added only when user clicks `Add`. |
| Holiday List (`holiday-list`) | Chip list (read-only + removable) | No | `No holiday dates selected.` | Displays selected holiday dates. | Dates are unique and sorted; each chip can be removed with `x`. |

#### Area: Actions

##### Fields

| Action | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Refresh (`refresh-btn`) | Button | No | Enabled | Reloads profiles from `GET /api/capacity/profiles`. | Shows error if API call fails. |
| New (`new-btn`) | Button | No | Enabled | Clears current selection and opens blank editor state. | Does not delete any saved profile. |
| Save (`save-btn`) | Button | No | Enabled | Saves current form to `POST /api/capacity`. | Upserts by `from_date + to_date` key. |
| Delete (`delete-btn`) | Button | No | Enabled when profiles exist | Deletes selected profile via `DELETE /api/capacity`. | Requires profile selection and user confirmation. |
| Add Holiday (`holiday-add`) | Button | No | Enabled | Adds currently selected date from calendar picker into holiday list. | Ignores empty selection; deduplicates values. |
| Clear Holidays (`holiday-clear`) | Button | No | Enabled | Removes all currently selected holiday dates from the form state. | Does not delete saved profiles until `Save`. |

### Status Area

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Status (`status`) | Read-only text | No | Empty | Displays operation feedback (loading, success, errors). | Class `ok`/`err` used for visual state. |

## Dialogs

### Delete Confirmation Dialog

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Confirm Delete | Browser confirm dialog | Yes | N/A | Confirms deletion of selected range. | Cancel keeps data unchanged. |

## Behavior Summary

- On load, the page calls `GET /api/capacity/profiles` and populates the dropdown.
- Selecting a profile copies its values into editor fields.
- Holiday dates are managed through calendar picker + chips (not comma text input).
- Saving sends normalized form payload to `POST /api/capacity`.
- Deleting sends selected `from/to` to `DELETE /api/capacity`.
