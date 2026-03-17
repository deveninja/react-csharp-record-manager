using server.Contracts;
using server.Data;

namespace server.Endpoints;

public static class RecordEndpoints
{
	public static IEndpointRouteBuilder MapRecordEndpoints(this IEndpointRouteBuilder endpoints)
	{
		var group = endpoints.MapGroup("/api/records").WithTags("Records");

		group.MapGet("/", (IRecordRepository repository) =>
			Results.Ok(repository.GetAll()));

		group.MapGet("/options", (IRecordRepository repository) =>
			Results.Ok(repository.GetOptions()));

		group.MapPost("/", (UpdateRecordRequest request, IRecordRepository repository) =>
		{
			var created = repository.TryCreate(request, out var createdRecord, out var validationError);
			if (!created)
			{
				return Results.BadRequest(validationError);
			}

			return Results.Created($"/api/records/{createdRecord.Id}", createdRecord);
		});

		group.MapGet("/{id:int}", (int id, IRecordRepository repository) =>
		{
			var record = repository.GetById(id);
			return record is null ? Results.NotFound() : Results.Ok(record);
		});

		group.MapDelete("/{id:int}", (int id, IRecordRepository repository) =>
		{
			var deleted = repository.Delete(id);
			return deleted ? Results.NoContent() : Results.NotFound();
		});

		group.MapPut("/{id:int}", (int id, UpdateRecordRequest request, IRecordRepository repository) =>
		{
			var updated = repository.TryUpdate(id, request, out var updatedRecord, out var validationError);
			if (!updated)
			{
				return Results.NotFound();
			}

			if (!string.IsNullOrWhiteSpace(validationError))
			{
				return Results.BadRequest(validationError);
			}

			return Results.Ok(updatedRecord);
		});

		return endpoints;
	}
}
