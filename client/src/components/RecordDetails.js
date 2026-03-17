import { formatOptionLabel } from "../utils/formatters";

function RecordDetails({
  selectedRecord,
  form,
  isEditing,
  isCreating,
  saving,
  categoryOptions,
  statusOptions,
  onChange,
  onCreate,
  onEdit,
  onSave,
  onSaveNew,
  onDelete,
  onCancel,
}) {
  const canEdit = isEditing || isCreating;

  return (
    <section className="detail-panel">
      <h2>Record Details</h2>
      {!selectedRecord && !isCreating ? (
        <>
          <p>Select a record to view details.</p>
          <button type="button" onClick={onCreate} disabled={saving}>
            Add New Record
          </button>
        </>
      ) : (
        <div className="form-grid">
          <p className="mode-indicator">
            Mode: {isCreating ? "Create" : isEditing ? "Edit" : "View"}
          </p>

          <label htmlFor="name">Name</label>
          <input
            id="name"
            name="name"
            value={form.name}
            onChange={onChange}
            disabled={!canEdit || saving}
          />

          <label htmlFor="category">Category</label>
          <select
            id="category"
            name="category"
            value={form.category}
            onChange={onChange}
            disabled={!canEdit || saving}
          >
            <option value="" disabled>
              Select category
            </option>
            {categoryOptions.map((option) => (
              <option key={option} value={option}>
                {formatOptionLabel(option)}
              </option>
            ))}
          </select>

          <label htmlFor="status">Status</label>
          <select
            id="status"
            name="status"
            value={form.status}
            onChange={onChange}
            disabled={!canEdit || saving}
          >
            <option value="" disabled>
              Select status
            </option>
            {statusOptions.map((option) => (
              <option key={option} value={option}>
                {formatOptionLabel(option)}
              </option>
            ))}
          </select>

          <label htmlFor="description">Description</label>
          <textarea
            id="description"
            name="description"
            value={form.description}
            onChange={onChange}
            rows="5"
            disabled={!canEdit || saving}
          />

          <div className="actions-row">
            {isCreating ? (
              <>
                <button type="button" onClick={onSaveNew} disabled={saving}>
                  {saving ? "Saving..." : "Create"}
                </button>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={onCancel}
                  disabled={saving}
                >
                  Cancel
                </button>
              </>
            ) : !isEditing ? (
              <>
                <button type="button" onClick={onEdit} disabled={saving}>
                  Edit
                </button>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={onDelete}
                  disabled={saving}
                >
                  Delete
                </button>
              </>
            ) : (
              <>
                <button type="button" onClick={onSave} disabled={saving}>
                  {saving ? "Saving..." : "Save"}
                </button>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={onCancel}
                  disabled={saving}
                >
                  Cancel
                </button>
              </>
            )}
          </div>
        </div>
      )}
    </section>
  );
}

export default RecordDetails;
