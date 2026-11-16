package com.takipkolay.mobile;

import android.Manifest;
import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;

final class ReminderReceiver extends BroadcastReceiver {
    private static final String CHANNEL = "takip_kolay_reminders";

    @Override
    public void onReceive(Context context, Intent intent) {
        if (intent == null || !ReminderScheduler.ACTION_REMIND.equals(intent.getAction())
                || intent.getData() == null) {
            return;
        }
        String id = intent.getData().getLastPathSegment();
        if (id == null) return;

        try {
            Task match = null;
            for (Task task : new TaskStore(context).load()) {
                if (id.equals(task.id)) {
                    match = task;
                    break;
                }
            }
            if (match == null || match.completed || match.reminderAt == null
                    || match.reminderAt > System.currentTimeMillis() + 60_000L) return;
        } catch (Exception ignored) {
            // Fail closed: never show a reminder for inaccessible or corrupt data.
            return;
        }
        if (Build.VERSION.SDK_INT >= 33
                && context.checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS)
                != PackageManager.PERMISSION_GRANTED) {
            return;
        }

        NotificationManager manager = (NotificationManager) context.getSystemService(Context.NOTIFICATION_SERVICE);
        if (manager == null) return;
        NotificationChannel channel = new NotificationChannel(CHANNEL, "Görev hatırlatmaları",
                NotificationManager.IMPORTANCE_DEFAULT);
        channel.setLockscreenVisibility(Notification.VISIBILITY_SECRET);
        manager.createNotificationChannel(channel);

        Intent open = new Intent(context, MainActivity.class);
        open.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_CLEAR_TOP);
        open.setData(Uri.parse("takipkolay://task/" + Uri.encode(id)));
        PendingIntent tap = PendingIntent.getActivity(context, id.hashCode(), open,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
        Notification notification = new Notification.Builder(context, CHANNEL)
                .setSmallIcon(android.R.drawable.ic_dialog_info)
                .setContentTitle("Takip Kolay")
                .setContentText("Takip zamanınız geldi. Görevlerinizi kontrol edin.")
                .setCategory(Notification.CATEGORY_REMINDER)
                .setVisibility(Notification.VISIBILITY_SECRET)
                .setContentIntent(tap)
                .setAutoCancel(true)
                .build();
        manager.notify(id.hashCode(), notification);
    }
}
