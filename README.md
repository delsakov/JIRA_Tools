# JIRA_Tools
This repository would contain some code for integration with JIRA.
- Note: Executable "compiled" files located in "distr"/ folder.
- Note 2: Please do not forget about licence - copyright is required.


# JIRA Migration Tool
This is Tool with UI for migrating one JIRA project on one instance to different JIRA instanse and different Instance.
Project admin rights are required.

Steps:
1. Create Mapping File (dynamicaly creating Excel sheets)
2. Provide the required mappings
3. Run the Tool with that file for migration process (select the objects to be migrated)
4. Wait

Note: All IDs are copyed as is (OLD_PROJECT-12345 will be NEW_PROJECT-12345), including Sprints, Issues, FixVersions, Components, Comments, Attachments and Links.

For any questions, bugs - please contact me (Dmitry Elsakov).

# JIRA Export Tool
Tool for exporting JIRA data to the Excel file. No limitation of number of issues. Including simple UI.
For all available fields - Project Admin is required OR only standard fields will be exported. - (TBD)

# JIRA Bulk Upload Tool (In Progress)
This Tool is for bulk-upload (import), bulk-update or bulk export.
Also could be used for creating snapshot of issues in Excel with ability to re-upload them later.
