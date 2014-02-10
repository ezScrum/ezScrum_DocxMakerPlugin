Ext.ns('Plugin.docxMaker.releasePlanToolBar');

Plugin.docxMaker.releasePlanToolBar.downloadBtnHandler = function(btn) {
	var selectedNode = Ext.getCmp('ReleasePlan_ReleaseTree').getSelectionModel().getSelectedNode();
	ModifyReleasePlanWindow.resetForm();
	var releaseID = ModifyReleasePlanWindow.loadData(selectedNode);
	var form = Ext.DomHelper.append(document.body, {
        tag : 'form',
        method : 'post',
        action : 'plugin/DocxMaker/getReleasePlan?releaseID=' + releaseID
    });
    document.body.appendChild(form);

    // add any other form fields you like here
    form.submit();
    document.body.removeChild(form); 

};


Plugin.docxMaker.releasePlanToolBar.downloadBtnPlugin = Ext.extend(Ext.util.Observable, {

	init: function(cmp) {
		this.hostCmp = cmp;
		this.hostCmp.on('render', this.onRender, this, {delay: 200});
	},

	onRender: function() {
		var downloadBtn = new Ext.Button({
			id: 'ReleasePlan_downloadReleaseBtn',
			disabled: true,
			text: 'Download Release',
			icon: 'pluginWorkspace/DocxMakerPlugin/resource/download.png',
			handler: Plugin.docxMaker.releasePlanToolBar.downloadBtnHandler
		});

		this.hostCmp.add(downloadBtn);
		this.hostCmp.doLayout();
	}
});

Ext.preg('downloadBtnPlugin', Plugin.docxMaker.releasePlanToolBar.downloadBtnPlugin);