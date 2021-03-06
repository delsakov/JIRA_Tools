// Created by Dmitry Elsakov
import com.atlassian.jira.component.ComponentAccessor
import com.onresolve.scriptrunner.canned.jira.admin.CopyProject
import groovy.transform.Field
import com.onresolve.scriptrunner.runner.rest.common.CustomEndpointDelegate
import groovy.json.JsonBuilder
import javax.ws.rs.core.Response
import groovy.transform.BaseScript
import javax.ws.rs.core.MultivaluedMap
import javax.servlet.http.HttpServletRequest

@BaseScript CustomEndpointDelegate delegate

createProject(httpMethod: "POST", group: ["u_jira_global_admin"]) { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def projectKey = request.getParameter("key")
    def projectName = request.getParameter("name")
    def projectParentKey = request.getParameter("parent")
if (projectKey) {
	
	def projectManager = ComponentAccessor.getProjectManager()
	def copyProject = new CopyProject()
    
	def parentProject = projectManager.getProjectByCurrentKey(projectParentKey)	
  	def params = [
      FIELD_SOURCE_PROJECT : parentProject.getKey(),
      FIELD_TARGET_PROJECT : projectKey.toString(),
      FIELD_TARGET_PROJECT_NAME : projectName.toString(),
      FIELD_COPY_VERSIONS : false,
      FIELD_COPY_COMPONENTS : false,
      FIELD_COPY_ISSUES : false,
      FIELD_COPY_DASH_AND_FILTERS : false
   	]
    try {
   		copyProject.doScript(params)
    } 
    catch (Exception ex) {
   
    	return Response.ok(new JsonBuilder("Error": ex.message).toString()).build()
    }
    
    return Response.ok(new JsonBuilder("Project created": projectKey).toString()).build()
}
else {
    def message = "'key' should be specified as params. Example: .../createProject?key=<NEW_KEY>&name=<NEW PROJECT NAME>"
    return Response.ok(new JsonBuilder("Error": message).toString()).build()
}
}
