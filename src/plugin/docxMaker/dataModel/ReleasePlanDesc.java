package plugin.docxMaker.dataModel;

import java.util.LinkedList;
import java.util.List;

public class ReleasePlanDesc {
	private String ID;
	private String Name;
	private String StartDate;
	private String EndDate;
	private String Description;
	private List<ISprintPlanDesc> SprintList;
	
	public ReleasePlanDesc() {
		this.ID = "0";
		this.Name = "";
		this.StartDate = "";
		this.EndDate = "";
		this.Description = "";
		this.SprintList = new LinkedList<ISprintPlanDesc>();
	}

	public void setID(String id) {
		this.ID = id;
	}
	
	public void setName(String Name) {
		this.Name = Name;
	}
	
	public void setStartDate(String StartDate) {
		this.StartDate = StartDate;		
	}
	
	public void setEndDate(String EndDate) {
		this.EndDate = EndDate;
	}
	
	public void setDescription(String Description) {
		this.Description = Description;
	}
	
	public String getID() {
		return this.ID;
	}
	
	public String getName() {
		return this.Name;
	}
	
	public String getStartDate() {
		return this.StartDate;
	}
	
	public String getEndDate() {
		return this.EndDate;
	}
	
	public String getDescription() {
		return this.Description;
	}

	public List<ISprintPlanDesc> getSprintDescList() {
		return this.SprintList;
	}

	public void setSprintDescList(List<ISprintPlanDesc> SprintList) {
		this.SprintList = SprintList;
	}
}
