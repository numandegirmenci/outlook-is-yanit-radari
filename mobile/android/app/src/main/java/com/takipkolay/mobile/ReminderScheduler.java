package com.takipkolay.mobile;

import android.app.AlarmManager;
import android.app.PendingIntent;
import android.content.Context;
import android.content.Intent;
import android.net.Uri;

final class ReminderScheduler {
    static final String ACTION_REMIND = "com.takipkolay.mobile.REMIND";
    private ReminderScheduler() {}

    private static PendingIntent pendingIntent(Context context, String taskId) {
        Intent intent = new Intent(context, ReminderReceiver.class);
        intent.setAction(ACTION_REMIND);
        intent.setData(Uri.parse("takipkolay://reminder/" + Uri.encode(taskId)));
        return PendingIntent.getBroadcast(context, 0, intent,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
    }

    static void cancel(Context context, String taskId) {
        AlarmManager alarms = (AlarmManager) context.getSystemService(Context.ALARM_SERVICE);
        if (alarms != null) {
            alarms.cancel(pendingIntent(context, taskId));
        }
    }

    static void schedule(Context context, Task task) {
        if (task.completed || task.reminderAt == null || task.reminderAt <= System.currentTimeMillis()) {
            return;
        }
        AlarmManager alarms = (AlarmManager) context.getSystemService(Context.ALARM_SERVICE);
        if (alarms != null) {
            // Inexact by design: requires no special exact-alarm permission.
            alarms.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, task.reminderAt,
                    pendingIntent(context, task.id));
        }
    }
}
