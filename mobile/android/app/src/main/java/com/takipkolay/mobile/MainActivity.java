package com.takipkolay.mobile;

import android.Manifest;
import android.app.Activity;
import android.app.AlertDialog;
import android.app.DatePickerDialog;
import android.app.TimePickerDialog;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.graphics.Typeface;
import android.os.Build;
import android.os.Bundle;
import android.text.InputType;
import android.view.View;
import android.view.WindowManager;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.Toast;

import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.UUID;

public final class MainActivity extends Activity {
    private static final int NOTIFICATION_REQUEST = 71;
    private static final int MAX_FIELD_LENGTH = 200_000;
    private final Locale tr = new Locale("tr", "TR");

    private TaskStore store;
    private ScrollView scroll;
    private EditText titleInput;
    private EditText noteInput;
    private EditText sourceInput;
    private Button reminderButton;
    private Button saveButton;
    private LinearLayout taskList;
    private Long reminderAt;
    private String editingId;

    @Override
    protected void onCreate(Bundle state) {
        super.onCreate(state);
        getWindow().addFlags(WindowManager.LayoutParams.FLAG_SECURE);
        store = new TaskStore(this);
        buildScreen();
        refreshTasks();
        acceptIncoming(getIntent());
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        setIntent(intent);
        acceptIncoming(intent);
    }

    private int dp(float value) {
        return (int) (getResources().getDisplayMetrics().density * value + 0.5f);
    }

    private LinearLayout column() {
        LinearLayout view = new LinearLayout(this);
        view.setOrientation(LinearLayout.VERTICAL);
        return view;
    }

    private void addSpace(LinearLayout parent, int dp) {
        View v = new View(this);
        parent.addView(v, new LinearLayout.LayoutParams(1, dp(dp)));
    }

    private TextView text(String value, int sp, boolean bold) {
        TextView v = new TextView(this);
        v.setText(value);
        v.setTextSize(sp);
        v.setTextColor(Color.rgb(35, 42, 54));
        if (bold) v.setTypeface(null, Typeface.BOLD);
        return v;
    }

    private Button button(String caption, LinearLayout parent) {
        Button b = new Button(this);
        b.setText(caption);
        b.setAllCaps(false);
        parent.addView(b, new LinearLayout.LayoutParams(
                LinearLayout.LayoutParams.MATCH_PARENT, LinearLayout.LayoutParams.WRAP_CONTENT));
        return b;
    }

    private EditText input(LinearLayout parent, String label, String hint, int lines) {
        parent.addView(text(label, 14, true));
        EditText field = new EditText(this);
        field.setTextSize(15);
        field.setHint(hint);
        field.setInputType(InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_FLAG_MULTI_LINE
                | InputType.TYPE_TEXT_FLAG_CAP_SENTENCES);
        field.setMinLines(lines);
        field.setMaxLines(lines == 1 ? 3 : 12);
        parent.addView(field, new LinearLayout.LayoutParams(
                LinearLayout.LayoutParams.MATCH_PARENT, LinearLayout.LayoutParams.WRAP_CONTENT));
        addSpace(parent, 8);
        return field;
    }

    private void buildScreen() {
        scroll = new ScrollView(this);
        scroll.setFillViewport(true);
        LinearLayout body = column();
        body.setPadding(dp(18), dp(22), dp(18), dp(30));
        scroll.addView(body);
        setContentView(scroll);

        body.addView(text("Takip Kolay", 26, true));
        body.addView(text("Outlook mobil için yerel görev takibi", 13, false));
        addSpace(body, 18);

        body.addView(text("Yeni görev / görev düzenle", 18, true));
        addSpace(body, 10);
        titleInput = input(body, "E-posta konusu / görev başlığı",
                "Örn. Mağaza ruhsat evrakları", 1);
        noteInput = input(body, "Takip notu",
                "Örn. Ruhsat evraklarının tamamlanmasını takip et.", 3);
        sourceInput = input(body, "E-posta / yazışma içeriği",
                "Outlook'tan paylaşılan veya yapıştırılan metin burada görünür.", 5);

        reminderButton = button("Hatırlatma: tarih seçilmedi", body);
        reminderButton.setOnClickListener(v -> chooseDateTime());
        Button clearDate = button("Hatırlatma tarihini kaldır (tarihsiz)", body);
        clearDate.setOnClickListener(v -> {
            reminderAt = null;
            renderReminderLabel();
        });
        addSpace(body, 8);
        saveButton = button("Görevi kaydet", body);
        saveButton.setOnClickListener(v -> saveTask());
        Button fresh = button("Yeni boş görev", body);
        fresh.setOnClickListener(v -> resetForm());
        addSpace(body, 8);
        body.addView(text("Bilgiler yalnızca bu uygulamanın özel, şifreli yerel dosyasında tutulur. "
                + "E-posta okunması veya Microsoft hesabına otomatik erişim yoktur.", 12, false));
        addSpace(body, 26);

        body.addView(text("Kayıtlı görevler", 19, true));
        body.addView(text("Açmak / düzenlemek için görev başlığına dokunun.", 12, false));
        addSpace(body, 12);
        taskList = column();
        body.addView(taskList);
        renderReminderLabel();
    }

    private void renderReminderLabel() {
        if (reminderAt == null) {
            reminderButton.setText("Hatırlatma: tarih seçilmedi");
        } else {
            DateFormat format = DateFormat.getDateTimeInstance(
                    DateFormat.MEDIUM, DateFormat.SHORT, tr);
            reminderButton.setText("Hatırlatma: " + format.format(new Date(reminderAt)));
        }
    }

    private void chooseDateTime() {
        Calendar initial = Calendar.getInstance();
        if (reminderAt != null) initial.setTimeInMillis(reminderAt);
        DatePickerDialog date = new DatePickerDialog(this, (picker, year, month, day) -> {
            TimePickerDialog time = new TimePickerDialog(this, (clock, hour, minute) -> {
                Calendar chosen = Calendar.getInstance();
                chosen.set(Calendar.YEAR, year);
                chosen.set(Calendar.MONTH, month);
                chosen.set(Calendar.DAY_OF_MONTH, day);
                chosen.set(Calendar.HOUR_OF_DAY, hour);
                chosen.set(Calendar.MINUTE, minute);
                chosen.set(Calendar.SECOND, 0);
                chosen.set(Calendar.MILLISECOND, 0);
                if (chosen.getTimeInMillis() <= System.currentTimeMillis()) {
                    toast("Gelecekteki bir tarih ve saat seçin.");
                    return;
                }
                reminderAt = chosen.getTimeInMillis();
                renderReminderLabel();
            }, initial.get(Calendar.HOUR_OF_DAY), initial.get(Calendar.MINUTE), true);
            time.setTitle("Hatırlatma saatini seç");
            time.show();
        }, initial.get(Calendar.YEAR), initial.get(Calendar.MONTH), initial.get(Calendar.DAY_OF_MONTH));
        date.setTitle("Hatırlatma tarihini seç");
        date.getDatePicker().setMinDate(System.currentTimeMillis() - 1_000L);
        date.show();
    }

    private void acceptIncoming(Intent intent) {
        if (intent == null) return;
        String action = intent.getAction();
        String shared = null;
        String subject = null;
        if (Intent.ACTION_SEND.equals(action) && "text/plain".equals(intent.getType())) {
            shared = intent.getStringExtra(Intent.EXTRA_TEXT);
            subject = intent.getStringExtra(Intent.EXTRA_SUBJECT);
        } else if (Intent.ACTION_PROCESS_TEXT.equals(action)) {
            CharSequence selection = intent.getCharSequenceExtra(Intent.EXTRA_PROCESS_TEXT);
            if (selection != null) shared = selection.toString();
        } else if (intent.getData() != null
                && "takipkolay".equals(intent.getData().getScheme())
                && "task".equals(intent.getData().getHost())) {
            openTask(intent.getData().getLastPathSegment());
            return;
        }

        if ((shared == null || shared.isEmpty()) && (subject == null || subject.isEmpty())) return;
        // The share target receives only what the source app explicitly exports.
        resetForm();
        if (subject != null) titleInput.setText(subject);
        if (shared != null) {
            if (shared.length() > MAX_FIELD_LENGTH) {
                new AlertDialog.Builder(this)
                        .setTitle("Paylaşılan metin çok uzun")
                        .setMessage("Metin 200.000 karakteri geçiyor. Otomatik kırpma yapılmadı. "
                                + "Daha kısa bir bölümü paylaşın.")
                        .setPositiveButton("Tamam", null).show();
            } else {
                sourceInput.setText(shared);
                if (shared.trim().startsWith("http") && !shared.contains("\n")) {
                    toast("Yalnızca bağlantı paylaşılmış olabilir. E-posta metnini kontrol edin.");
                }
            }
        }
        scroll.smoothScrollTo(0, 0);
    }

    private void saveTask() {
        String title = titleInput.getText().toString().trim();
        String note = noteInput.getText().toString().trim();
        String source = sourceInput.getText().toString().trim();
        if (title.isEmpty()) {
            titleInput.setError("Görev başlığı gerekli.");
            titleInput.requestFocus();
            return;
        }
        if (title.length() > 500 || note.length() > MAX_FIELD_LENGTH
                || source.length() > MAX_FIELD_LENGTH) {
            toast("Başlık veya metin çok uzun. Bilgi otomatik kesilmedi.");
            return;
        }
        try {
            List<Task> tasks = store.load();
            Task task = null;
            if (editingId != null) {
                for (Task item : tasks) {
                    if (editingId.equals(item.id)) {
                        task = item;
                        break;
                    }
                }
                if (task == null) {
                    toast("Güncellenecek görev bulunamadı.");
                    return;
                }
            }
            if (reminderAt != null && reminderAt <= System.currentTimeMillis()
                    && (task == null || !reminderAt.equals(task.reminderAt))) {
                toast("Yeni hatırlatma gelecekte olmalı. Tarihi değiştirin veya kaldırın.");
                return;
            }
            if (task == null) {
                task = new Task();
                task.id = UUID.randomUUID().toString();
                task.createdAt = System.currentTimeMillis();
                tasks.add(task);
            }
            task.title = title;
            task.note = note;
            task.sourceText = source;
            task.reminderAt = reminderAt;
            store.save(tasks);
            ReminderScheduler.cancel(this, task.id);
            ReminderScheduler.schedule(this, task);
            if (task.reminderAt != null && !task.completed) requestNotificationPermissionIfNeeded();
            resetForm();
            refreshTasks();
            toast("Görev telefona kaydedildi.");
        } catch (Exception ex) {
            showStorageError();
        }
    }

    private void requestNotificationPermissionIfNeeded() {
        if (Build.VERSION.SDK_INT >= 33
                && checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS)
                != PackageManager.PERMISSION_GRANTED) {
            requestPermissions(new String[]{Manifest.permission.POST_NOTIFICATIONS},
                    NOTIFICATION_REQUEST);
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String[] permissions, int[] grants) {
        super.onRequestPermissionsResult(requestCode, permissions, grants);
        if (requestCode == NOTIFICATION_REQUEST
                && (grants.length == 0 || grants[0] != PackageManager.PERMISSION_GRANTED)) {
            toast("Görev kaydedildi, fakat bildirim izni olmadan telefon uyarı gösteremez.");
        }
    }

    private void refreshTasks() {
        taskList.removeAllViews();
        List<Task> tasks;
        try {
            tasks = new ArrayList<>(store.load());
        } catch (Exception ex) {
            showStorageError();
            return;
        }
        tasks.sort(Comparator.comparing((Task t) -> t.completed)
                .thenComparingLong(t -> t.reminderAt == null ? Long.MAX_VALUE : t.reminderAt)
                .thenComparing((Task t) -> t.createdAt, Comparator.reverseOrder()));
        if (tasks.isEmpty()) {
            taskList.addView(text("Henüz görev yok. Yukarıdan ilk görevi ekleyebilirsiniz.", 14, false));
            return;
        }

        DateFormat format = DateFormat.getDateTimeInstance(
                DateFormat.MEDIUM, DateFormat.SHORT, tr);
        for (Task task : tasks) {
            LinearLayout card = column();
            card.setPadding(dp(12), dp(12), dp(12), dp(12));
            card.setBackgroundColor(Color.rgb(241, 245, 250));
            TextView heading = text((task.completed ? "✓ " : "• ") + task.title, 16, true);
            card.addView(heading);
            String info = task.completed ? "Tamamlandı" :
                    task.reminderAt == null ? "Tarihsiz takip" :
                            "Hatırlatma: " + format.format(new Date(task.reminderAt));
            card.addView(text(info, 13, false));
            String preview = !task.note.isEmpty() ? task.note : task.sourceText;
            if (!preview.isEmpty()) {
                preview = preview.replace("\n", " ");
                if (preview.length() > 160) preview = preview.substring(0, 160) + "…";
                card.addView(text(preview, 13, false));
            }

            Button open = button("Aç / düzenle", card);
            open.setOnClickListener(v -> openTask(task.id));
            Button toggle = button(task.completed ? "Yeniden aç" : "Tamamlandı", card);
            toggle.setOnClickListener(v -> toggleCompleted(task.id));
            Button delete = button("Sil", card);
            delete.setOnClickListener(v -> new AlertDialog.Builder(this)
                    .setTitle("Görev silinsin mi?")
                    .setMessage("Bu kayıt yalnızca cihazda tutuluyor. Silme geri alınamaz.")
                    .setNegativeButton("İptal", null)
                    .setPositiveButton("Sil", (dialog, which) -> deleteTask(task.id))
                    .show());
            taskList.addView(card, new LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT, LinearLayout.LayoutParams.WRAP_CONTENT));
            addSpace(taskList, 10);
        }
    }

    private void openTask(String id) {
        if (id == null) return;
        try {
            for (Task task : store.load()) {
                if (!id.equals(task.id)) continue;
                editingId = task.id;
                titleInput.setText(task.title);
                noteInput.setText(task.note);
                sourceInput.setText(task.sourceText);
                reminderAt = task.reminderAt;
                renderReminderLabel();
                saveButton.setText("Görevi güncelle");
                scroll.smoothScrollTo(0, 0);
                return;
            }
            toast("Görev bulunamadı.");
        } catch (Exception ex) {
            showStorageError();
        }
    }

    private void toggleCompleted(String id) {
        try {
            List<Task> tasks = store.load();
            for (Task task : tasks) {
                if (!task.id.equals(id)) continue;
                task.completed = !task.completed;
                store.save(tasks);
                ReminderScheduler.cancel(this, task.id);
                ReminderScheduler.schedule(this, task);
                refreshTasks();
                return;
            }
        } catch (Exception ex) {
            showStorageError();
        }
    }

    private void deleteTask(String id) {
        try {
            List<Task> tasks = store.load();
            boolean removed = tasks.removeIf(task -> task.id.equals(id));
            if (!removed) return;
            store.save(tasks);
            ReminderScheduler.cancel(this, id);
            if (id.equals(editingId)) resetForm();
            refreshTasks();
            toast("Görev silindi.");
        } catch (Exception ex) {
            showStorageError();
        }
    }

    private void resetForm() {
        editingId = null;
        reminderAt = null;
        titleInput.setText("");
        noteInput.setText("");
        sourceInput.setText("");
        saveButton.setText("Görevi kaydet");
        renderReminderLabel();
    }

    private void showStorageError() {
        new AlertDialog.Builder(this)
                .setTitle("Şifreli kayıt açılamadı")
                .setMessage("Görev dosyası okunamadı veya kaydedilemedi. Mevcut veriler "
                        + "otomatik sıfırlanmadı. Uygulamayı kaldırmadan önce sorunu inceleyin.")
                .setPositiveButton("Tamam", null)
                .show();
    }

    private void toast(String value) {
        Toast.makeText(this, value, Toast.LENGTH_LONG).show();
    }
}
