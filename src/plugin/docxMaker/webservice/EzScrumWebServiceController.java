package plugin.docxMaker.webservice;

import java.io.File;

import javax.ws.rs.core.MediaType;

import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;

import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.WebResource;
import com.sun.jersey.api.client.WebResource.Builder;

import plugin.docxMaker.DocxMaker;

public class EzScrumWebServiceController {
	private String mEzScrumURL;

	public EzScrumWebServiceController(String ezScrumURL) {
		mEzScrumURL = ezScrumURL;
	}

	/**
	 * use web-service to get release info and return release plan file
	 * 
	 * @param encodedUserName - the user account(encoded)
	 * @param encodedPassword - the user password(encoded)
	 * @param projectId - the release of project
	 * @param releaseId - the release docx you wnat
	 * @return release docx file
	 * @throws JSONException
	 */
	public File getReleaseDocx(String encodedUserName, String encodedPassword, String projectId, String releaseId) throws JSONException {
		StringBuilder releaseWebServiceUrl = new StringBuilder();
		releaseWebServiceUrl.append("http://")
					        .append(mEzScrumURL)
					        .append("/web-service/").append(projectId)
					        .append("/release-plan/").append(releaseId)
					        .append("/all").append("?username=").append(encodedUserName).append("&password=").append(encodedPassword);

		Client client = Client.create();
		WebResource webResource = client.resource(releaseWebServiceUrl.toString());
		Builder result = webResource.type(MediaType.APPLICATION_JSON).accept(MediaType.APPLICATION_JSON);
		JSONObject releaseJSON = result.get(JSONObject.class);
		return new DocxMaker().getReleasePlanDocx(releaseJSON);
	}
}
