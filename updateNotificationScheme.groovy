// Created by Dmitry Elsakov
import com.atlassian.jira.permission.PermissionSchemeManager
import com.atlassian.jira.scheme.SchemeEntity
import com.atlassian.jira.notification.NotificationSchemeManager
import com.atlassian.jira.component.ComponentAccessor
import groovy.transform.Field
import com.onresolve.scriptrunner.runner.rest.common.CustomEndpointDelegate
import groovy.json.JsonBuilder
import javax.ws.rs.core.Response
import groovy.transform.BaseScript
import javax.ws.rs.core.MultivaluedMap
import javax.servlet.http.HttpServletRequest

@BaseScript CustomEndpointDelegate delegate

updateNotificationScheme(httpMethod: "POST", group: ["u_jira_global_admin"]) { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def projectKey = request.getParameter("key")
    def notificationSchemeId = request.getParameter("schemeId")
if (projectKey && notificationSchemeId) {
    
	NotificationSchemeManager nsm = ComponentAccessor.getNotificationSchemeManager();
	def notificationScheme = nsm.getSchemeObject(notificationSchemeId.toInteger());
	
    def projectManager = ComponentAccessor.getProjectManager()
	def project = projectManager.getProjectByCurrentKey(projectKey)	
	
    try {
        nsm.removeSchemesFromProject(project);
    	nsm.addSchemeToProject(project, notificationScheme);
	} 
    catch (Exception ex) {
   		return Response.status(404).entity(new JsonBuilder("Error": ex.message).toString()).build()
    }
    
    return Response.ok(new JsonBuilder("Notification Scheme ID updated to:": notificationSchemeId).toString()).build()
}
else {
    def message = "'key' and 'schemeId' should be specified as params. Example: .../updateNotificationScheme?key=<PROJECT_KEY>&schemeId=<NOTIFICATION_SCHEME_ID>"
    return Response.status(401).entity(new JsonBuilder("Error": message).toString()).build()
	}
}
