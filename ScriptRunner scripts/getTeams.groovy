// Created by Dmitry Elsakov
import com.onresolve.scriptrunner.runner.customisers.WithPlugin
import com.onresolve.scriptrunner.runner.customisers.PluginModule
import com.atlassian.jira.component.ComponentAccessor
import com.atlassian.rm.teams.api.team.GeneralTeamService  
@WithPlugin("com.atlassian.teams")
@PluginModule GeneralTeamService teamService 
import com.onresolve.scriptrunner.runner.rest.common.CustomEndpointDelegate
import groovy.json.JsonBuilder
import groovy.transform.BaseScript
import javax.ws.rs.core.MultivaluedMap
import javax.servlet.http.HttpServletRequest
import javax.ws.rs.core.Response
import org.codehaus.jackson.map.ObjectMapper

def customFieldManager = ComponentAccessor.getCustomFieldManager()

@BaseScript CustomEndpointDelegate delegate

getTeams(httpMethod: "GET") { MultivaluedMap queryParams, body, HttpServletRequest request ->
    def teamIds = request.getParameter("teamIds")
if (teamIds) {
    def number_params = teamIds.toString().split(',')
    def teamNames = []
    
    for (def i=0; i<number_params.size(); i++) {
        try {
			def teamName = teamService.getTeam(number_params[i].toLong())?.get().description?.title
    		teamNames << (["id":number_params[i], "name": teamName])
        }
        catch (Exception e1) {
            teamNames << (["id":number_params[i], "name": "Unknown"])
        }
    }
        return Response.ok(new JsonBuilder("teamName": teamNames).toString()).build()
    }
 
else {
    def teamMap = teamService.teamIdsWithoutPermissionCheck.collectEntries {[(it)]} 
    def teamNames = []
        
    teamMap.each { key, value ->
    try {
			def teamNameAll = teamService.getTeam(key.toLong())?.get().description?.title
    		teamNames << (["id":key, "name": teamNameAll])
        }
        catch (Exception e1) {
            teamNames << (["id":key, "name": "Unknown"])
        }
    }
        return Response.ok(new JsonBuilder("teamName": teamNames).toString()).build()
    }
}
