package com.takipkolay.mobile;

import org.json.JSONException;
import org.json.JSONObject;

final class Task {
    String id;
    String title;
    String note;
    String sourceText;
    long createdAt;
    Long reminderAt;
    boolean completed;

    JSONObject toJson() throws JSONException {
        JSONObject item = new JSONObject();
        item.put("id", id);
        item.put("title", title);
        item.put("note", note);
        item.put("sourceText", sourceText);
        item.put("createdAt", createdAt);
        item.put("completed", completed);
        item.put("reminderAt", reminderAt == null ? JSONObject.NULL : reminderAt);
        return item;
    }

    static Task fromJson(JSONObject json) throws JSONException {
        Task task = new Task();
        task.id = json.getString("id");
        task.title = json.getString("title");
        task.note = json.optString("note", "");
        task.sourceText = json.optString("sourceText", "");
        task.createdAt = json.optLong("createdAt", 0L);
        task.completed = json.optBoolean("completed", false);
        task.reminderAt = json.isNull("reminderAt") ? null : json.getLong("reminderAt");
        return task;
    }
}
