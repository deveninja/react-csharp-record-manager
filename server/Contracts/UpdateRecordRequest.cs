namespace server.Contracts;

public record UpdateRecordRequest(
	string Name,
	string Category,
	string Status,
	string Description);
