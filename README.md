# React + C# Record Manager

Full-stack sample app with:

- UI: React (Create React App) in `client`
- API: ASP.NET Core (.NET 9) in `server`

## 1. Prerequisites

Install these tools first:

- Node.js 18+ (or newer LTS)
- npm (comes with Node.js)
- .NET SDK 9.0+

Quick checks:

```bash
node -v
npm -v
dotnet --version
```

## 2. Project Structure

- `client`: React UI
- `server`: ASP.NET Core API
- `server/server.http`: ready-to-run HTTP requests for API testing

## 3. Environment Configuration

### 3.1 UI env file

The UI reads API base URL from:

- `client/.env`

Current key:

```env
REACT_APP_API_ROOT=http://localhost:5000/api
```

If needed, copy from `client/.env.example`.

### 3.2 API settings (C# equivalent to .env)

For this API, runtime values are centralized in:

- `server/Properties/launchSettings.json` (local run profiles)
- `server/appsettings.json`
- `server/appsettings.Development.json`

Current important values:

- `ASPNETCORE_URLS=http://localhost:5000`
- `Client__Origin=http://localhost:3000` (CORS origin)

Note: ASP.NET Core does not automatically load a `.env` file by default.

## 4. Install Dependencies

From repo root:

```bash
cd client
npm install
```

No extra install step is needed for server beyond .NET restore/build during run.

## 5. Run the API

Open terminal 1:

```bash
cd server
dotnet run
```

Expected API base URL:

- `http://localhost:5000`

## 6. Run the UI

Open terminal 2:

```bash
cd client
npm start
```

Expected UI URL:

- `http://localhost:3000`

## 7. Verify the App

1. Open `http://localhost:3000`
2. Confirm records list loads from API
3. Try these flows:
   - Select record and edit/save
   - Add new record
   - Delete record

## 8. Verify API Endpoints Directly

Use `server/server.http` in VS Code REST Client style, or use curl/Postman.

Available routes:

- `GET /api/records/`
- `GET /api/records/options`
- `GET /api/records/{id}`
- `POST /api/records/`
- `PUT /api/records/{id}`
- `DELETE /api/records/{id}`

## 9. Build Commands

### UI production build

```bash
cd client
npm run build
```

### API build

```bash
cd server
dotnet build
```

## 10. Common Issues

### UI cannot reach API

Check:

- API is running on `http://localhost:5000`
- `client/.env` has `REACT_APP_API_ROOT=http://localhost:5000/api`
- CORS origin in launch settings is `http://localhost:3000`

### Port conflicts

If port `3000` or `5000` is busy:

- Update UI/API ports in their configs
- Keep these in sync:
  - `client/.env` (`REACT_APP_API_ROOT`)
  - `server/Properties/launchSettings.json` (`ASPNETCORE_URLS`, `Client__Origin`)

## 11. Git Ignore Notes

Root `.gitignore` is configured for both folders:

- React artifacts (`client/node_modules`, `client/build`, etc.)
- .NET artifacts (`server/bin`, `server/obj`, etc.)
