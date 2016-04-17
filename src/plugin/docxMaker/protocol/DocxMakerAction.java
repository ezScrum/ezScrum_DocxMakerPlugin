package plugin.docxMaker.protocol;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;

import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;

import org.codehaus.jettison.json.JSONException;
import org.kohsuke.stapler.StaplerRequest;
import org.kohsuke.stapler.StaplerResponse;

import ntut.csie.protocal.Action;
import plugin.docxMaker.webservice.EzScrumWebServiceController;

public class DocxMakerAction implements Action {
	@Override
	public String getUrlName() {
		return "DocxMaker";
	}

	public void doGetReleasePlan(StaplerRequest request, StaplerResponse response) {
		HttpSession session = request.getSession();
		String encodedPassword = (String) session.getAttribute("passwordForPlugin");
		String projectName = getURLParameter(request, "projectName");
		String releaseId = request.getParameter("releaseID");
		Cookie[] cookies = request.getCookies();
		String encodedUserName = cookies[1].getValue();
		InputStream in = null;
		PrintWriter writer = null;
		try {
			String ezScrumURL = request.getServerName() + ":" + request.getLocalPort() + request.getContextPath();
			EzScrumWebServiceController service = new EzScrumWebServiceController(ezScrumURL);
			File releasePlan = service.getReleaseDocx(encodedUserName, encodedPassword, projectName, releaseId);
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
	
	private static String getURLParameter(HttpServletRequest request, String paramName) {
		try {
			/**
			 * 拿到request header的URL parameter 如果拿不到referer資訊此時回傳null
			 */
			String referer = request.getHeader("Referer");
			if (referer == null) return null;
			// 將URL轉換成沒有萬用字元符號
			String url = URLDecoder.decode(referer, "UTF-8");

			/**
			 * 如果用 "?"拆除出來的params是null或只有一個urlParams 代表網址後面沒有帶參數 此時回傳null
			 */
			String[] urlParams = url.split("\\?");
			if (urlParams == null || urlParams.length <= 1) return null;

			// 取得URL 問號後面的Parameter參數
			url = urlParams[1];
			// params 裡面存的大概像這樣 [{id=15}, {op=55}]，為各個parameter
			String[] params = url.split("&");
			// value 用來接params的 parameter的key跟value ex: [id, 15]
			String[] value;

			/**
			 * 找到對應的parameter name並把value回傳給使用者
			 */
			for (int i = 0; i < params.length; i++) {
				value = params[i].split("=");
				if (value[0].equals(paramName)) {
					return value[1];
				}
			}
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return paramName;
	}
}
