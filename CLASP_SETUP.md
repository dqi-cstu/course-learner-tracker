# v2.56.2 — Deploy / Update

## Existing tracker
1. Replace/push the project files with this release (`clasp push`).
2. In Apps Script choose **Deploy > Manage deployments > Edit**.
3. Select **New version** and deploy.
4. Keep the existing execution/access settings.
5. **Do not run `setupProject()` again.** Schema remains 14.

## New tracker
Create a new bound Google Sheet project, push the files, deploy the Web App, open it, and use the first-run Web initialization flow. Manual `setupProject()` is normally not needed.

## Section management
Go to **Course > วัน เวลา และสถานที่เรียน**. Use the **⋮** menu on an individual Section to edit or delete only that Section. Use **+ เพิ่ม Section** for a new Section.
