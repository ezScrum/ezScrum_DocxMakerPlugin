package plugin.docxMaker.webservice;

import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

import javax.ws.rs.core.MediaType;

import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;

import plugin.docxMaker.DocxMaker;
import plugin.docxMaker.dataModel.ReleasePlanObject;
import plugin.docxMaker.dataModel.SprintPlanObject;
import plugin.docxMaker.dataModel.StoryObject;
import plugin.docxMaker.dataModel.TaskObject;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.WebResource;
import com.sun.jersey.api.client.WebResource.Builder;
import com.sun.jersey.core.util.Base64;

public class EzScrumWebServiceController {
	private String mEzScrumURL;

	public EzScrumWebServiceController(String ezScrumURL) {
		mEzScrumURL = ezScrumURL;
	}

	public File getReleaseDocx(String account, String encodedPassword, String projectId, String releaseId) throws JSONException {
		String encodedUserName = new String(Base64.encode(account.getBytes()));
		// 需要的帳密為暗碼
		StringBuilder releaseWebServiceUrl = new StringBuilder();
		releaseWebServiceUrl.append("http://")
					        .append(mEzScrumURL)
					        .append("/web-service/").append(projectId)
					        .append("/release-plan/").append(releaseId)
					        .append("/all/").append("?userName=").append(encodedUserName).append("&password=").append(encodedPassword);

		Client client = Client.create();
		WebResource webResource = client.resource(releaseWebServiceUrl.toString());
		Builder result = webResource.type(MediaType.APPLICATION_JSON).accept(MediaType.APPLICATION_JSON);
		JSONObject resultJSONObject = result.get(JSONObject.class);
		Gson gson = new Gson();
		ReleasePlanObject reDesc = gson.fromJson(resultJSONObject.getString("releasePlanDesc"), ReleasePlanObject.class);
		List<SprintPlanObject> sprinDescList = gson.fromJson(resultJSONObject.getString("sprintDescList"), new TypeToken<List<SprintPlanObject>>() {}.getType());
		HashMap<String, List<StoryObject>> stories = gson.fromJson(resultJSONObject.getString("stories"), new TypeToken<HashMap<String, List<StoryObject>>>() {}.getType());
		LinkedHashMap<String, List<TaskObject>> taskMap = gson.fromJson(resultJSONObject.getString("taskMap"), new TypeToken<LinkedHashMap<String, List<TaskObject>>>() {}.getType());
		HashMap<String, Float> tatolStoryPoints = gson.fromJson(resultJSONObject.getString("totalStoryPoints"), new TypeToken<HashMap<String, Float>>() {}.getType());

		return new DocxMaker().getReleasePlanDocx(reDesc, sprinDescList, stories, taskMap, tatolStoryPoints);
	}
}
