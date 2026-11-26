# Security Specification — Portal E-Guru v10.2

## 1. Authentication & Authorization Model
- **Authentication**: Firebase Authentication (Google Sign-In) with mandatory `email_verified == true`. No plaintext passwords are stored in Firestore or client state.
- **Bootstrap Admin**: `rifairwansani@sma1kudus.sch.id` (always recognized as `Admin`).
- **Staff Authorization (`isStaff()`)**: Either `isBootstrapAdmin()` OR an active record in `/accounts/{emailLower}` (`status == 'Aktif'`).
- **Admin Authorization (`isAdmin()`)**: Either `isBootstrapAdmin()` OR an active record in `/accounts/{emailLower}` with `role == 'Admin'`.
- **Unregistered Users**: Signed-in users who are neither Bootstrap Admin nor listed as `Aktif` in `/accounts` can only create/read their own `/users/{uid}` document (with `role == 'Pending'`) and check their own `/accounts/{authEmail}` document (`get` only). They cannot read or write any school operational data.

## 2. Collection Access Matrix
| Collection | Get | List | Create | Update | Delete |
| --- | --- | --- | --- | --- | --- |
| `/users/{userId}` | Owner or Admin | Admin | Owner (`role == 'Pending'` or Bootstrap Admin `'Admin'`) | Admin, or Owner (non-role fields, `createdAt` immutable) | Admin |
| `/accounts/{accountId}` | Self (`authEmail() == accountId`) or Staff | Staff | Admin (`accountId == email.lower()`) | Admin, or Staff self updating `lastLogin` only | Admin (cannot delete Bootstrap Admin) |
| `/students/{studentId}` | Staff | Staff | Staff | Staff (`id` immutable) | Staff |
| `/teachers/{teacherId}` | Staff | Staff | Staff | Staff (`id` immutable) | Staff |
| `/schedules/{scheduleId}` | Staff | Staff | Staff | Staff (`id` immutable) | Staff |
| `/journals/{journalId}` | Staff | Staff | Staff | Staff (`id` & `createdAt` immutable) | Staff |
| `/grades/{gradeId}` | Staff | Staff | Staff | Staff (`id` immutable) | Staff |
| `/settings/global` | Staff | Staff | Staff | Staff | Admin |
| Catch-all `/{document=**}` | Deny | Deny | Deny | Deny | Deny |

## 3. Data Integrity & Anti-Abuse Constraints
- All IDs validated (`1..160` chars, no `/`).
- String lengths bounded on every field (`<= 250,000` chars on image DataURLs to stay well below Firestore's 1MB document limit).
- Numeric grade scores constrained to `null` or `0..100`.
- Strict `keys().hasAll(...)` and `keys().hasOnly(...)` field whitelisting on all entities.
- Server timestamps (`request.time`) enforced on `serverCreatedAt` (create-only, immutable on update) and `serverUpdatedAt` on writes.
