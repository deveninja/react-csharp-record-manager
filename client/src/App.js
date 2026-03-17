import "./App.css";
import { useEffect, useMemo, useState } from "react";
import { fetchJsonWithCache, invalidateCache } from "./utils/apiCache";
import { formatOptionLabel } from "./utils/formatters";
import RecordsList from "./components/RecordsList";
import RecordDetails from "./components/RecordDetails";

const API_ROOT = process.env.REACT_APP_API_ROOT || "http://localhost:5000/api";
const RECORDS_URL = `${API_ROOT}/records`;
const OPTIONS_URL = `${API_ROOT}/records/options`;
const RECORDS_TTL_MS = 30_000;
const OPTIONS_TTL_MS = 5 * 60_000;

const emptyForm = {
  name: "",
  category: "",
  status: "",
  description: "",
};

function App() {
  const [records, setRecords] = useState([]);
  const [categoryOptions, setCategoryOptions] = useState([]);
  const [statusOptions, setStatusOptions] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [form, setForm] = useState(emptyForm);
  const [isEditing, setIsEditing] = useState(false);
  const [isCreating, setIsCreating] = useState(false);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    const loadRecords = async () => {
      try {
        setLoading(true);
        setError("");

        const [recordsData, optionsData] = await Promise.all([
          fetchJsonWithCache(RECORDS_URL, { ttlMs: RECORDS_TTL_MS }),
          fetchJsonWithCache(OPTIONS_URL, { ttlMs: OPTIONS_TTL_MS }),
        ]);

        setRecords(recordsData);
        setCategoryOptions(optionsData.categories || []);
        setStatusOptions(optionsData.statuses || []);
      } catch (loadError) {
        const message = loadError.message.includes(OPTIONS_URL)
          ? "Unable to load record options."
          : "Unable to load records.";
        setError(message);
      } finally {
        setLoading(false);
      }
    };

    loadRecords();
  }, []);

  const selectedRecord = useMemo(
    () => records.find((record) => record.id === selectedId) || null,
    [records, selectedId],
  );

  const statusCounts = useMemo(() => {
    return records.reduce((acc, record) => {
      acc[record.status] = (acc[record.status] || 0) + 1;
      return acc;
    }, {});
  }, [records]);

  const selectedCount = selectedRecord ? 1 : 0;

  const handleRowSelect = (record) => {
    setSelectedId(record.id);
    setIsEditing(false);
    setIsCreating(false);
    setForm({
      name: record.name,
      category: record.category,
      status: record.status,
      description: record.description,
    });
  };

  const handleChange = (event) => {
    if (!isEditing && !isCreating) {
      return;
    }

    const { name, value } = event.target;
    setForm((prev) => ({ ...prev, [name]: value }));
  };

  const handleEdit = () => {
    if (!selectedRecord || isCreating) {
      return;
    }

    setIsEditing(true);
  };

  const handleCreate = () => {
    setSelectedId(null);
    setIsEditing(false);
    setIsCreating(true);
    setError("");
    setForm({
      name: "",
      category: categoryOptions[0] || "",
      status: statusOptions[0] || "",
      description: "",
    });
  };

  const handleCancel = () => {
    if (isCreating) {
      setIsCreating(false);
      setForm(emptyForm);
      return;
    }

    if (!selectedRecord) {
      return;
    }

    setIsEditing(false);
    setForm({
      name: selectedRecord.name,
      category: selectedRecord.category,
      status: selectedRecord.status,
      description: selectedRecord.description,
    });
  };

  const handleSave = async () => {
    if (!selectedRecord) {
      return;
    }

    try {
      setSaving(true);
      setError("");

      const response = await fetch(`${RECORDS_URL}/${selectedRecord.id}`, {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(form),
      });

      if (!response.ok) {
        throw new Error("Unable to save record.");
      }

      const updatedRecord = await response.json();
      setRecords((prevRecords) =>
        prevRecords.map((record) =>
          record.id === updatedRecord.id ? updatedRecord : record,
        ),
      );
      invalidateCache(RECORDS_URL);

      setForm({
        name: updatedRecord.name,
        category: updatedRecord.category,
        status: updatedRecord.status,
        description: updatedRecord.description,
      });
      setIsEditing(false);
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  };

  const handleSaveNew = async () => {
    if (!isCreating) {
      return;
    }

    try {
      setSaving(true);
      setError("");

      const response = await fetch(RECORDS_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(form),
      });

      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "Unable to add record.");
      }

      const createdRecord = await response.json();
      setRecords((prevRecords) => [...prevRecords, createdRecord]);
      invalidateCache(RECORDS_URL);

      setSelectedId(createdRecord.id);
      setForm({
        name: createdRecord.name,
        category: createdRecord.category,
        status: createdRecord.status,
        description: createdRecord.description,
      });
      setIsCreating(false);
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!selectedRecord || isCreating) {
      return;
    }

    const confirmed = window.confirm(
      `Delete record "${selectedRecord.name}"? This cannot be undone.`,
    );

    if (!confirmed) {
      return;
    }

    try {
      setSaving(true);
      setError("");

      const response = await fetch(`${RECORDS_URL}/${selectedRecord.id}`, {
        method: "DELETE",
      });

      if (!response.ok) {
        throw new Error("Unable to delete record.");
      }

      setRecords((prevRecords) =>
        prevRecords.filter((record) => record.id !== selectedRecord.id),
      );
      invalidateCache(RECORDS_URL);

      setSelectedId(null);
      setIsEditing(false);
      setIsCreating(false);
      setForm(emptyForm);
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="app-shell">
      <h1>Record Manager</h1>

      <section className="summary-panel">
        <p>Total records: {records.length}</p>
        <p>Selected records: {selectedCount}</p>
        <p>
          By status:{" "}
          {Object.keys(statusCounts).length === 0
            ? "None"
            : Object.entries(statusCounts)
                .map(
                  ([status, count]) =>
                    `${formatOptionLabel(status)} (${count})`,
                )
                .join(", ")}
        </p>
        <button
          type="button"
          onClick={handleCreate}
          disabled={saving || loading}
        >
          Add New Record
        </button>
      </section>

      {error && <p className="error">{error}</p>}

      <div className="layout-grid">
        <RecordsList
          records={records}
          selectedId={selectedId}
          loading={loading}
          onSelectRecord={handleRowSelect}
        />

        <RecordDetails
          selectedRecord={selectedRecord}
          form={form}
          isEditing={isEditing}
          isCreating={isCreating}
          saving={saving}
          categoryOptions={categoryOptions}
          statusOptions={statusOptions}
          onChange={handleChange}
          onCreate={handleCreate}
          onEdit={handleEdit}
          onSave={handleSave}
          onSaveNew={handleSaveNew}
          onDelete={handleDelete}
          onCancel={handleCancel}
        />
      </div>
    </div>
  );
}

export default App;
