import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class jFrame extends JFrame {

    public void cFrame() throws IOException{


        //添加窗口组件
        Panel p1 = new Panel();
        p1.setLayout(null);
        p1.setBounds(0,0,200,40);

        //添加文本框
        final JTextField jt1 = new JTextField();
        jt1.setEditable(false);
        jt1.setBounds(400,50,500,50);
        Font f1 = new Font("Helvetica",Font.PLAIN,18);
        jt1.setFont(f1);
        jt1.setText("D:\\A\\b");

        //添加textArea
        Font f2 = new Font("Helvetica",Font.PLAIN,22);
        final JTextArea jta1 = new JTextArea();
        jta1.setFont(f2);
        jta1.setBounds(50,300,850,400);
        jta1.setBackground(Color.black);
        jta1.setEditable(false);
        jta1.setForeground(Color.cyan);

        //添加按钮1
        final JButton button1  = new JButton("选择目录");
        button1.setBounds(50,50,300,50);
        button1.setFont(f1);
        button1.setFocusPainted(false);

        //添加按钮2
        final JButton button2  = new JButton("开始合并");
        button2.setBounds(50,150,300,100);
        button2.setFont(f1);
        button2.setFocusPainted(false);

        //内部类为按钮button1绑定事件监控器
        button1.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                //打开文件夹选择窗口
                jt1.setText(null);

                JFileChooser jfc=new JFileChooser(".");
                jfc.setFileSelectionMode(1);
                int state=jfc.showOpenDialog(null);//此句是打开文件选择器界面的触发语句
                if(state==1){
                    return;//撤销则返回
                }
                else{
                    File f=jfc.getSelectedFile();//f为选择到的目录
                    String dir = f.getAbsolutePath();
                    jt1.setText(dir);
                    button2.setEnabled(true);
                }
            }
        });

        //内部类为按钮button2绑定事件监控器

        button2.addActionListener(new ActionListener() {

            public void actionPerformed(ActionEvent e) {
                //创建输出流
                fileO fo = new fileO();
                fo.createFile("test");
                //开始合并文件
                String s = jt1.getText();
                if (s.length() <= 0){
                    JOptionPane.showMessageDialog(null, "亲,你忘了选择目录了吧!!!", "Warn!!!", JOptionPane.ERROR_MESSAGE);
                }
                List<String> fileList = new ArrayList<String>(POIUtil.getFilesName(s));
                button2.setEnabled(false);

                for (int i = 0; i < fileList.size(); i++) {
                    String fileName = fileList.get(i);
                    System.out.println(fileName);
                    try {

                        List<String[]> data = POIUtil.readExcel(fileName);
                        for (int j = 0; j < data.size(); j++) {
                            Object line = data.get(j);
                            String line2 = Arrays.toString((Object[]) line);
                            String line3 = line2.replace("[","");
                            String line4 = line3.replace("]","");
                            System.out.println(line4);
                            fo.writeFileContent("E:\\bao\\merge-excel\\merge-excel-all\\merge-excel-all.csv",line4);
                        }
                    } catch (IOException e1) {
                        e1.printStackTrace();
                    }
                }

                jt1.setText("文件合并完成啦(●ˇ∀ˇ●)");
            }
        });

        p1.add(jta1);
        p1.add(button1);
        p1.add(button2);
        p1.add(jt1);

        this.add(p1);

        //设置窗体title
        this.setTitle("合并Excel文件  BY  mhf  v1.4  O(∩_∩)O");
        //设置窗体大小
        this.setBounds(300,100,1200,800);
        //设置窗体关闭方式
        this.setDefaultCloseOperation(jFrame.EXIT_ON_CLOSE);
        //设置显示窗体
        this.setVisible(true);

    }

    public static void main(String[] args) {
        JFrame jFrame = new jFrame();
        try {
            ((jFrame) jFrame).cFrame();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
