package com.takipkolay.mobile;

import android.content.Context;
import android.security.keystore.KeyGenParameterSpec;
import android.security.keystore.KeyProperties;
import android.util.AtomicFile;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.security.KeyStore;
import java.util.ArrayList;
import java.util.List;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.crypto.spec.GCMParameterSpec;

final class TaskStore {
    private static final String ALIAS = "takip_kolay_tasks_aes_v1";
    private static final byte[] MAGIC = new byte[]{'T', 'K', '0', '1'};
    private static final int IV_SIZE = 12;
    private static final int MAX_FILE_BYTES = 32 * 1024 * 1024;
    private final AtomicFile file;

    TaskStore(Context context) {
        file = new AtomicFile(new File(context.getFilesDir(), "tasks.enc"));
    }

    synchronized List<Task> load() throws Exception {
        byte[] bytes;
        try (FileInputStream input = file.openRead()) {
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            byte[] buffer = new byte[8192];
            int n;
            while ((n = input.read(buffer)) != -1) {
                if (output.size() + n > MAX_FILE_BYTES) {
                    throw new IllegalStateException("Görev verisi beklenenden büyük.");
                }
                output.write(buffer, 0, n);
            }
            bytes = output.toByteArray();
        } catch (FileNotFoundException notFound) {
            return new ArrayList<>();
        }

        if (bytes.length < MAGIC.length + IV_SIZE + 16) {
            throw new IllegalStateException("Görev dosyası eksik.");
        }
        for (int i = 0; i < MAGIC.length; i++) {
            if (bytes[i] != MAGIC[i]) {
                throw new IllegalStateException("Görev dosyası tanınmadı.");
            }
        }

        byte[] iv = new byte[IV_SIZE];
        System.arraycopy(bytes, MAGIC.length, iv, 0, IV_SIZE);
        Cipher cipher = Cipher.getInstance("AES/GCM/NoPadding");
        cipher.init(Cipher.DECRYPT_MODE, getKey(), new GCMParameterSpec(128, iv));
        byte[] plain = cipher.doFinal(bytes, MAGIC.length + IV_SIZE, bytes.length - MAGIC.length - IV_SIZE);
        JSONObject root = new JSONObject(new String(plain, StandardCharsets.UTF_8));
        if (root.getInt("version") != 1) {
            throw new IllegalStateException("Desteklenmeyen veri sürümü.");
        }
        JSONArray items = root.getJSONArray("tasks");
        List<Task> tasks = new ArrayList<>();
        for (int i = 0; i < items.length(); i++) {
            tasks.add(Task.fromJson(items.getJSONObject(i)));
        }
        return tasks;
    }

    synchronized void save(List<Task> tasks) throws Exception {
        JSONObject root = new JSONObject();
        root.put("version", 1);
        JSONArray items = new JSONArray();
        for (Task task : tasks) {
            items.put(task.toJson());
        }
        root.put("tasks", items);

        byte[] plain = root.toString().getBytes(StandardCharsets.UTF_8);
        if (plain.length > MAX_FILE_BYTES - 32) {
            throw new IllegalStateException("Yerel görev depolama sınırına ulaşıldı.");
        }
        Cipher cipher = Cipher.getInstance("AES/GCM/NoPadding");
        cipher.init(Cipher.ENCRYPT_MODE, getKey());
        byte[] iv = cipher.getIV();
        if (iv.length != IV_SIZE) {
            throw new IllegalStateException("Beklenmeyen şifreleme biçimi.");
        }
        byte[] encrypted = cipher.doFinal(plain);
        FileOutputStream output = null;
        try {
            output = file.startWrite();
            output.write(MAGIC);
            output.write(iv);
            output.write(encrypted);
            file.finishWrite(output);
        } catch (Exception e) {
            if (output != null) {
                file.failWrite(output);
            }
            throw e;
        }
    }

    private SecretKey getKey() throws Exception {
        KeyStore store = KeyStore.getInstance("AndroidKeyStore");
        store.load(null);
        if (store.containsAlias(ALIAS)) {
            return (SecretKey) store.getKey(ALIAS, null);
        }
        KeyGenerator generator = KeyGenerator.getInstance(KeyProperties.KEY_ALGORITHM_AES, "AndroidKeyStore");
        generator.init(new KeyGenParameterSpec.Builder(
                ALIAS, KeyProperties.PURPOSE_ENCRYPT | KeyProperties.PURPOSE_DECRYPT)
                .setBlockModes(KeyProperties.BLOCK_MODE_GCM)
                .setEncryptionPaddings(KeyProperties.ENCRYPTION_PADDING_NONE)
                .setKeySize(256)
                .build());
        return generator.generateKey();
    }
}
