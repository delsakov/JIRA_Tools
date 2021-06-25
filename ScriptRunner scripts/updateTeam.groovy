// Created by Dmitry Elsakov
import com.atlassian.jira.issue.IssueManager
import com.atlassian.jira.component.ComponentAccessor
import com.atlassian.jira.issue.ModifiedValue
import com.atlassian.jira.issue.util.DefaultIssueChangeHolder
import com.atlassian.jira.issue.Issue
import groovy.transform.Field
import com.onresolve.scriptrunner.runner.rest.common.CustomEndpointDelegate
import groovy.json.JsonBuilder
import javax.ws.rs.core.Response
import groovy.transform.BaseScript
import javax.ws.rs.core.MultivaluedMap
import javax.servlet.http.HttpServletRequest

@BaseScript CustomEndpointDelegate delegate

updateTeam(httpMethod: "POST", group: ["u_jira_global_admin"]) { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def issueKey = request.getParameter("key")
    def customFieldId = request.getParameter("fieldId")
    def newTeamId = request.getParameter("newId")
if (issueKey && customFieldId) {
	
	def issue = ComponentAccessor.issueManager.getIssueByCurrentKey(issueKey)
	def customFieldManager = ComponentAccessor.getCustomFieldManager()
    def TeamField = customFieldManager.getCustomFieldObject(customFieldId)
    def TeamFieldType = TeamField.getCustomFieldType()  
    
    try {
        TeamField.updateValue(null, issue, new ModifiedValue(issue.getCustomFieldValue(TeamField), null), new DefaultIssueChangeHolder())
        if (newTeamId) {
   			TeamField.updateValue(null, issue, new ModifiedValue(null, TeamFieldType.getSingularObjectFromString(newTeamId)), new DefaultIssueChangeHolder())
        }
    } 
    catch (Exception ex) {
   		return Response.status(404).entity(new JsonBuilder("Error": ex.message).toString()).build()
    }
    
    return Response.ok(new JsonBuilder("Team updated to:": newTeamId).toString()).build()
}
else {
    def message = "'key' and 'fieldId' should be specified as params; 'newId' is optional. Example: .../updateTeam?key=<ISSUE_KEY>&fieldId=<CUSTOM_FIELD_ID for TEAM>&newID=<NEW_TEAM_ID>"
    return Response.status(401).entity(new JsonBuilder("Error": message).toString()).build()
	}
}
