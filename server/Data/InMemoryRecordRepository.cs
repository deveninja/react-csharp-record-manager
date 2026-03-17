using server.Contracts;
using server.Domain;

namespace server.Data;

public sealed class InMemoryRecordRepository : IRecordRepository
{
	private readonly List<RecordItem> _records =
	[
		new(1, "Website Redesign", CategoryType.Project, RecordStatus.InProgress, "Refresh the marketing site and improve mobile layout."),
		new(2, "Northwind Account", CategoryType.Customer, RecordStatus.Active, "Primary enterprise customer in the retail vertical."),
		new(3, "Alicia Porter", CategoryType.Employee, RecordStatus.OnLeave, "Senior designer currently on approved leave."),
		new(4, "Q2 Security Audit", CategoryType.Task, RecordStatus.Pending, "Complete policy review and vulnerability follow-up."),
		new(5, "Onboarding Revamp", CategoryType.Project, RecordStatus.Completed, "Updated onboarding checklist and training assets."),
		new(6, "Partner Onboarding", CategoryType.Customer, RecordStatus.Active, "Coordinate implementation milestones and training plan."),
		new(7, "Device Lifecycle Policy", CategoryType.Task, RecordStatus.InProgress, "Finalize replacement cadence and ownership matrix."),
		new(8, "Hiring Loop Calibration", CategoryType.Employee, RecordStatus.Pending, "Align interview scorecard expectations across panel."),
		new(9, "Q3 Planning", CategoryType.Project, RecordStatus.Pending, "Create staffing and dependency forecast for next quarter."),
		new(10, "Content Migration", CategoryType.Project, RecordStatus.InProgress, "Move legacy docs to the new portal taxonomy."),
		new(11, "SOC2 Evidence Review", CategoryType.Task, RecordStatus.Active, "Collect and validate control evidence with security."),
		new(12, "Design System Adoption", CategoryType.Project, RecordStatus.Completed, "Rolled out shared components across product areas."),
	];

	public IReadOnlyList<RecordItem> GetAll() => _records;

	public RecordItem? GetById(int id) => _records.FirstOrDefault(r => r.Id == id);

	public RecordOptionsResponse GetOptions() =>
		new(Enum.GetNames<CategoryType>(), Enum.GetNames<RecordStatus>());

	public bool Delete(int id)
	{
		var index = _records.FindIndex(r => r.Id == id);
		if (index < 0)
		{
			return false;
		}

		_records.RemoveAt(index);
		return true;
	}

	public bool TryCreate(UpdateRecordRequest request, out RecordItem createdRecord, out string? validationError)
	{
		createdRecord = default!;
		validationError = null;

		if (!TryParseRequest(request, out var category, out var status, out validationError))
		{
			return false;
		}

		var nextId = _records.Count == 0 ? 1 : _records.Max(r => r.Id) + 1;
		createdRecord = new RecordItem(
			nextId,
			request.Name,
			category,
			status,
			request.Description);

		_records.Add(createdRecord);
		return true;
	}

	public bool TryUpdate(int id, UpdateRecordRequest request, out RecordItem updatedRecord, out string? validationError)
	{
		updatedRecord = default!;
		validationError = null;

		var index = _records.FindIndex(r => r.Id == id);
		if (index < 0)
		{
			return false;
		}

		if (!TryParseRequest(request, out var category, out var status, out validationError))
		{
			return true;
		}

		updatedRecord = _records[index] with
		{
			Name = request.Name,
			Category = category,
			Status = status,
			Description = request.Description,
		};

		_records[index] = updatedRecord;
		return true;
	}

	private static bool TryParseRequest(
		UpdateRecordRequest request,
		out CategoryType category,
		out RecordStatus status,
		out string? validationError)
	{
		validationError = null;

		if (string.IsNullOrWhiteSpace(request.Name))
		{
			category = default;
			status = default;
			validationError = "Name is required.";
			return false;
		}

		if (!Enum.TryParse<CategoryType>(request.Category, ignoreCase: true, out category))
		{
			status = default;
			validationError = $"Invalid category '{request.Category}'.";
			return false;
		}

		if (!Enum.TryParse<RecordStatus>(request.Status, ignoreCase: true, out status))
		{
			validationError = $"Invalid status '{request.Status}'.";
			return false;
		}

		return true;
	}
}
