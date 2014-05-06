package plugin.docxMaker.protocol;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;

import javax.servlet.http.HttpSession;

import ntut.csie.ezScrum.pic.core.IUserSession;
import ntut.csie.ezScrum.web.form.ProjectInfoForm;
import ntut.csie.ezScrum.web.internal.IProjectSummaryEnum;
import ntut.csie.protocal.Action;

import org.codehaus.jettison.json.JSONException;
import org.kohsuke.stapler.StaplerRequest;
import org.kohsuke.stapler.StaplerResponse;

import plugin.docxMaker.webservice.EzScrumWebServiceController;

public class DocxMakerAction implements Action {
	@Override
	public String getUrlName() {
		return "DocxMaker";
	}

	public void doGetReleasePlan(StaplerRequest request, StaplerResponse response) {
		HttpSession session = request.getSession();
		ProjectInfoForm projectInfoForm = (ProjectInfoForm) session.getAttribute(IProjectSummaryEnum.PROJECT_INFO_FORM);
		IUserSession userSession = (IUserSession) session.getAttribute("UserSession");
		String userName = userSession.getAccount().getAccount();
		String encodedPassword = (String) session.getAttribute("passwordForPlugin");
		String projectId = projectInfoForm.getName();
		String releaseId = request.getParameter("releaseID");
		InputStream in = null;
		PrintWriter writer = null;
		try {
			String ezScrumURL = request.getServerName() + ":" + request.getLocalPort() + request.getContextPath();
			EzScrumWebServiceController service = new EzScrumWebServiceController(ezScrumURL);
			File releasePlan = service.getReleaseDocx(userName, encodedPassword, projectId, releaseId);
			response.setHeader("Content-Transfer-Encoding", "binary");
			response.setHeader("Content-Disposition", "attachment; filename=" + releasePlan.getName());
			response.setContentLength((int) releasePlan.length());
			response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
			writer = response.getWriter();
			in = new FileInputStream(releasePlan);
			int byteCode;
			while ((byteCode = in.read()) != -1)
				response.getWriter().write(byteCode);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (JSONException e) {
	        e.printStackTrace();
        } finally {
			try {
				if (in != null) in.close();
				if (writer != null) writer.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
