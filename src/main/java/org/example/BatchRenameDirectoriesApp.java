package org.example;

import javafx.application.Application;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import javafx.scene.control.Button;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class BatchRenameDirectoriesApp extends Application {

    private String excelFilePath; // Excel文件路径
    private String rootDirectory; // 原始目录路径

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage stage) throws Exception {
        BorderPane borderPane = new BorderPane();
        VBox vt=new VBox();
        VBox vc=new VBox();
        borderPane.setTop(vt);
        borderPane.setCenter(vc);
        stage.setWidth(701);
        stage.setHeight(701);
        stage.setScene(new Scene(borderPane));
        stage.setTitle("Rename Directories");

        // 创建ImageView并加载图片
        Image image = new Image("D:\\DAppSapce\\Code\\JavaProject\\XmlProject\\ChangeDirectoryNmae\\src\\main\\resources\\12logo.png"); // 替换为你的图片路径
        ImageView imageView = new ImageView(image);
        imageView.setFitWidth(300); // 设置图片宽度
        imageView.setPreserveRatio(true); // 保持图片宽高比例
        Label label=new Label("Rename Directories App");
        label.setStyle("-fx-font-family: 'Times New Roman'; -fx-font-size: 30px;");
        // 将ImageView添加到VBox中
        vt.getChildren().add(imageView);
        vt.getChildren().add(label);
        vt.setAlignment(Pos.CENTER); // 设置按钮在VBox中居中对齐
        // 创建按钮
        Button setExcelPathButton = new Button("选择Excel文件路径");
        setExcelPathButton.setStyle(
                "-fx-background-color: #B3EE3A;" + // 设置背景颜色为蓝色
                        "-fx-font-size: 24px;" + // 设置字体大小为24px
                        "-fx-pref-width: 380px;" + // 设置按钮的宽度
                        "-fx-pref-height: 80px;"+ // 设置按钮的高度
                        "-fx-background-radius: 20px;" // 设置按钮的圆角半径
        );
        Button setRootDirectoryButton = new Button("选择文件夹路径");
        setRootDirectoryButton.setStyle(
                "-fx-background-color: #B3EE3A;" + // 设置背景颜色为蓝色
                        "-fx-font-size: 24px;" + // 设置字体大小为24px
                        "-fx-pref-width: 380px;" + // 设置按钮的宽度
                        "-fx-pref-height: 80px;"+ // 设置按钮的高度
                        "-fx-background-radius: 20px;" // 设置按钮的圆角半径
        );
        setExcelPathButton.setOnAction(event -> {
            // 打开文件选择器，让用户选择Excel文件
            FileChooser fileChooser = new FileChooser();
            File excelFile = fileChooser.showOpenDialog(stage);
            if (excelFile != null) {
                excelFilePath = excelFile.getAbsolutePath();

                // 读取Excel文件内容，获取新旧名称映射关系
                Map<String, String> nameTable = readNameTableFromExcel(excelFilePath);

                // 调用目录重命名方法
                renameDirectories(rootDirectory, nameTable);
            }
        });

        setRootDirectoryButton.setOnAction(event -> {
            // 打开目录选择器，让用户选择原始目录
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File rootDir = directoryChooser.showDialog(stage);
            if (rootDir != null) {
                rootDirectory = rootDir.getAbsolutePath();

                // 读取Excel文件内容，获取新旧名称映射关系
                Map<String, String> nameTable = readNameTableFromExcel(excelFilePath);

                // 调用目录重命名方法
                renameDirectories(rootDirectory, nameTable);
            }

            // 显示转换完成的对话框
            Alert alert = new Alert(AlertType.INFORMATION);
            alert.setTitle("转换完成");
            alert.setHeaderText(null);
            alert.setContentText("目录重命名已完成！");
            alert.showAndWait();
        });


        // 将按钮添加到布局中
        vc.setSpacing(30); // 设置元素之间的垂直间距为10像素
        vc.setAlignment(Pos.CENTER); // 设置按钮在VBox中居中对齐
        vc.getChildren().addAll(setExcelPathButton, setRootDirectoryButton);
        stage.show();
    }

    private Map<String, String> readNameTableFromExcel(String excelFilePath) {
        Map<String, String> nameTable = new HashMap<>();

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFilePath))) {
            Sheet sheet = workbook.getSheetAt(0); // 假设名称在第一个工作表中

            for (Row row : sheet) {
                Cell oldNameCell = row.getCell(0);
                Cell newNameCell = row.getCell(1);

                if (oldNameCell != null && newNameCell != null) {
                    String oldName = oldNameCell.getStringCellValue().trim();
                    String newName = newNameCell.getStringCellValue().trim();

                    nameTable.put(oldName, newName);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return nameTable;
    }

    private void renameDirectories(String directoryPath, Map<String, String> nameTable) {
        File directory = new File(directoryPath);

        // 检查当前路径是否为目录
        if (!directory.isDirectory()) {
            return;
        }

        // 获取该目录下的所有子文件夹和文件
        File[] files = directory.listFiles();

        // 遍历子文件夹和文件
        for (File file : files) {
            // 如果是子目录，则递归调用该方法修改目录名称
            if (file.isDirectory()) {
                renameDirectories(file.getAbsolutePath(), nameTable);
            }

            // 修改当前目录的名称
            String oldName = file.getName();
            if (nameTable.containsKey(oldName)) {
                String newName = nameTable.get(oldName);
                File newDirectory = new File(file.getParent(), newName);
                file.renameTo(newDirectory);
            }
        }
    }
}
