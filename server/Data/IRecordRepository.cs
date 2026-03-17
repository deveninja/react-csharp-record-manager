using server.Contracts;
using server.Domain;

namespace server.Data;

public interface IRecordRepository
{
	IReadOnlyList<RecordItem> GetAll();
	RecordItem? GetById(int id);
	RecordOptionsResponse GetOptions();
	bool Delete(int id);
	bool TryCreate(UpdateRecordRequest request, out RecordItem createdRecord, out string? validationError);
	bool TryUpdate(int id, UpdateRecordRequest request, out RecordItem updatedRecord, out string? validationError);
}
