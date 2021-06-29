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

clearChangeHistory(httpMethod: "POST", group: ["jira-administrators"]) { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def issueKey = request.getParameter("key")

if (issueKey) {
	
	def issue = ComponentAccessor.issueManager.getIssueByCurrentKey(issueKey)
	def customFieldManager = ComponentAccessor.getCustomFieldManager()
	def changeHistoryManager = ComponentAccessor.getChangeHistoryManager()
    
    try {
        changeHistoryManager.removeAllChangeItems(issue)
    } 
    catch (Exception ex) {
   		return Response.status(404).entity(new JsonBuilder("Error": ex.message).toString()).build()
    }
    
    return Response.ok(new JsonBuilder("Change History cleared for: ": issueKey).toString()).build()
}
else {
    def message = "'key' should be specified as param. Example: .../clearChangeHistory?key=<ISSUE_KEY>"
    return Response.status(401).entity(new JsonBuilder("Error": message).toString()).build()
	}
}
