package com.takipkolay.mobile;

import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;

final class BootReceiver extends BroadcastReceiver {
    @Override
    public void onReceive(Context context, Intent intent) {
        if (intent == null) return;
        String action = intent.getAction();
        if (!Intent.ACTION_BOOT_COMPLETED.equals(action)
                && !Intent.ACTION_MY_PACKAGE_REPLACED.equals(action)) return;
        final PendingResult pending = goAsync();
        final Context app = context.getApplicationContext();
        new Thread(() -> {
            try {
                for (Task task : new TaskStore(app).load()) {
                    ReminderScheduler.schedule(app, task);
                }
            } catch (Exception ignored) {
                // Do not delete existing data when the keystore is unavailable.
            } finally {
                pending.finish();
            }
        }, "takip-kolay-alarm-restore").start();
    }
}
