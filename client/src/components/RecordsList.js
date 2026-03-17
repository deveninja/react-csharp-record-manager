function RecordsList({ records, selectedId, loading, onSelectRecord }) {
  return (
    <section className="list-panel">
      <h2>Records</h2>
      {loading ? (
        <p>Loading records...</p>
      ) : (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Name</th>
                <th>Category</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {records.map((record) => (
                <tr
                  key={record.id}
                  className={record.id === selectedId ? "selected-row" : ""}
                  onClick={() => onSelectRecord(record)}
                >
                  <td>{record.name}</td>
                  <td>{record.category}</td>
                  <td>{record.status}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

export default RecordsList;
