package minichat;

import java.util.ArrayList;
import java.util.List;

public class ChatGroup {
	private String groupName;
    private List<String> members;

    public ChatGroup(String groupName) {
        this.groupName = groupName;
        this.members = new ArrayList<>();
    }

    public String getGroupName() {
        return groupName;
    }

    public List<String> getMembers() {
        return members;
    }

    public void addMember(String username) {
        if (!members.contains(username)) {
            members.add(username);
        }
    }

    public boolean hasMember(String username) {
        return members.contains(username);
    }
}
