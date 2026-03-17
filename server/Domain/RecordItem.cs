namespace server.Domain;

public record RecordItem(
	int Id,
	string Name,
	CategoryType Category,
	RecordStatus Status,
	string Description);
