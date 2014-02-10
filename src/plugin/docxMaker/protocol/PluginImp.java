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
		final PluginUI pluginUI = new PluginUI() {
			public String getPluginID() {
				return "DocxMakerPlugin";
			}
		};
		ezScrumUIList.add(pluginUI);
		
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
