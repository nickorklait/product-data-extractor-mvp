# Secure customer pilot

Use the Databricks App URL for customer testing. Do not expose the local development server or create a public tunnel.

## 1. Deploy the production target

1. Deploy the latest `main` branch to the existing custom Databricks App.
2. Keep the app's organization setting at **Only people with access can use**.
3. Confirm `GET /api/health` returns `{"status":"ok"}`.
4. Confirm `GET /api/storage-status` returns `{"enabled":true}`.

## 2. Isolate stored pilot data

1. Create or select a dedicated managed Unity Catalog Volume for this pilot.
2. Add it to the app as a **UC volume** resource with key `volume`.
3. Grant the app resource **Can read and write** only on this Volume.
4. Do not grant the app access to unrelated catalogs, schemas, tables, warehouses, or model endpoints.
5. Limit direct Volume access to the pilot administrators and the app's dedicated service principal.

The app temporarily processes files on local app storage, deletes the temporary upload after extraction, and archives source documents and generated Excel files in the configured Volume.

## 3. Invite testers

1. Create a small account group such as `product-data-extractor-pilot-testers`, or select named users.
2. Open the app overview and select **Share**.
3. Add the group or users with **CAN USE**.
4. Keep **CAN MANAGE** limited to the app owner and one backup administrator.
5. Send testers the Databricks App URL. They must sign in through the organization's identity provider.

Databricks Apps do not support anonymous public access. External customer users must first be onboarded to the Databricks account through the approved identity-provider, SCIM, or JIT process.

## 4. Agree on pilot handling rules

Before real documents are uploaded, agree with the customer on:

- which documents are permitted in the pilot;
- who may access the source files and exports;
- the retention period for `source`, `failed`, and `excel` folders;
- who deletes pilot data when the test ends;
- where testers report the processing reference shown by an error.

Start with synthetic or non-sensitive documents. Use real confidential specifications only after the data owner approves the Volume location, access list, and retention period.

## 5. Acceptance check

Run this with a tester account that has only **CAN USE**:

1. Open the app and confirm sign-in is required.
2. Upload one valid PDF or DOCX smaller than 25 MB.
3. Confirm extracted values and source text appear.
4. Export the Phase 1 comparison workbook.
5. Confirm the source and Excel files appear in the dedicated Volume.
6. Confirm the tester cannot edit the app, its resources, or its permissions.
7. Remove the tester's access and confirm the app can no longer be opened.

## 6. Retention and closeout

The MVP does not automatically delete archived Volume files. During the pilot, assign an owner to remove files on the agreed schedule. At pilot closeout, remove tester access, delete retained test data as agreed, and review app and Volume audit events.
