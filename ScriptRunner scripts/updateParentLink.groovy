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

updateParentLink(httpMethod: "POST", group: ["u_jira_global_admin"]) { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def issueKey = request.getParameter("key")
    def customFieldId = request.getParameter("fieldId")
    def newParentKey = request.getParameter("parent")
if (issueKey && customFieldId) {
	
	def issue = ComponentAccessor.issueManager.getIssueByCurrentKey(issueKey)
	def customFieldManager = ComponentAccessor.getCustomFieldManager()
    def ParentLink = customFieldManager.getCustomFieldObject(customFieldId)
    def parentLinkFieldType = ParentLink.getCustomFieldType()  
    
    try {
        ParentLink.updateValue(null, issue, new ModifiedValue(issue.getCustomFieldValue(ParentLink), null), new DefaultIssueChangeHolder())
        if (newParentKey) {
            def parent = ComponentAccessor.issueManager.getIssueByCurrentKey(newParentKey)
   			ParentLink.updateValue(null, issue, new ModifiedValue(null, parentLinkFieldType.getSingularObjectFromString(parent.key)), new DefaultIssueChangeHolder())
        }
        
    } 
    catch (Exception ex) {
   		return Response.status(404).entity(new JsonBuilder("Error": ex.message).toString()).build()
    }
    
    return Response.ok(new JsonBuilder("Parent Link updated to:": newParentKey).toString()).build()
}
else {
    def message = "'key', 'fieldId' and 'parent' should be specified as params. Example: .../updateParentLink?key=<ISSUE_KEY>&fieldId=<CUSTOM_FIELD_ID for PARENT_LINK>&parent=<NEW_PARENT_ISSUE_KEY>"
    return Response.status(401).entity(new JsonBuilder("Error": message).toString()).build()
}
}
