// Created by Dmitry Elsakov
import com.atlassian.jira.component.ComponentAccessor
import com.atlassian.jira.scheme.SchemeManager
import com.atlassian.jira.notification.NotificationSchemeManager
import com.onresolve.scriptrunner.runner.rest.common.CustomEndpointDelegate
import groovy.json.JsonBuilder
import javax.ws.rs.core.Response
import groovy.transform.BaseScript
import javax.ws.rs.core.MultivaluedMap
import javax.servlet.http.HttpServletRequest

@BaseScript CustomEndpointDelegate delegate

setNotificationSchemeForProject(httpMethod: "POST", group: ["u_jira_global_admin"]) { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def projectKey = request.getParameter("key")
    def notificationSchemeName = request.getParameter("schemeName")

if (projectKey && notificationSchemeName) {
    
    def projectManager = ComponentAccessor.getProjectManager();
    def project = projectManager.getProjectByCurrentKey(projectKey);
    def notificationSchemeManager = ComponentAccessor.getNotificationSchemeManager();
    def notificationScheme = notificationSchemeManager.getSchemeObject(notificationSchemeName);
	
   try {
        notificationSchemeManager.removeSchemesFromProject(project);
   	notificationSchemeManager.addSchemeToProject(project, notificationScheme);
    } 
    catch (Exception ex) {
   	return Response.status(404).entity(new JsonBuilder("Error": ex.message).toString()).build()
    }
    
    return Response.ok(new JsonBuilder("Notification Scheme has been successfully updated.": projectKey).toString()).build()
	
}
else {
    def message = "'key' should be specified as param. Example: .../setNotificationSchemeForProject?key=<TARGET_PROJECT_KEY>&schemeName=<NOTIFICATION_SCHEME_NAME>"
    return Response.status(401).entity(new JsonBuilder("Error": message).toString()).build()
}
}
