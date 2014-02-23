package plugin.docxMaker.protocol;

import java.util.ArrayList;
import java.util.List;

import ntut.csie.ui.protocol.EzScrumUI;
import ntut.csie.ui.protocol.PluginUI;
import ntut.csie.ui.protocol.ReleasePlanUI;
import ntut.csie.ui.protocol.UIConfig;

public class PluginImp extends UIConfig {
	@Override
	public void setEzScrumUIList(List<EzScrumUI> ezScrumUIList) {
		/**
		 * add PluginUI to  ezScrumUIList for DoD Plug-in
		 */
		final PluginUI pluginUI = new PluginUI() {
			public String getPluginID() {
				return "DocxMakerPlugin";
			}
		};
		ezScrumUIList.add(pluginUI);
		
		/**
		 * add ReleasePlanUI to ezScrumUIList for ReleasePlan Pages view
		 */
		ReleasePlanUI releasePlanUI = new ReleasePlanUI() {
			@Override
			public List<String> getToolbarPluginIDList() {
				List<String> toolbarPluginIDList = new ArrayList<String>();
				toolbarPluginIDList.add("downloadBtnPlugin");
				return toolbarPluginIDList;
			}

			@Override
			public PluginUI getPluginUI() {
				return pluginUI;
			}
		};
		ezScrumUIList.add(releasePlanUI);
	}
}
