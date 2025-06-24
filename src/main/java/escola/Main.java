package escola;

import com.formdev.flatlaf.FlatLightLaf;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.awt.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.sql.*;
import java.util.Locale;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

public class Main {
    private static Connection connection;
    private static JFrame frame;
    private static JPanel mainPanel;
    private static CardLayout cardLayout;
    private static DefaultListModel<String> salasListModel;

    public static void main(String[] args) {
        try {
            FlatLightLaf.setup();
            UIManager.put("Button.arc", 10);
            UIManager.put("Component.arc", 10);
            UIManager.put("TextComponent.arc", 5);
            UIManager.put("ProgressBar.arc", 10);
            UIManager.put("ScrollBar.width", 12);
        } catch (Exception ex) {
            System.err.println("Failed to initialize LaF");
        }

        try {
            connection = DriverManager.getConnection("jdbc:sqlite:escola.db");
            criarTabelas();
        } catch (SQLException e) {
            JOptionPane.showMessageDialog(null, "Erro ao conectar ao banco de dados: " + e.getMessage());
            System.exit(1);
        }

        frame = new JFrame("Sistema Escolar");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(900, 650);
        frame.setLocationRelativeTo(null);

        cardLayout = new CardLayout();
        mainPanel = new JPanel(cardLayout);
        mainPanel.setBorder(new EmptyBorder(15, 15, 15, 15));
        mainPanel.setBackground(new Color(245, 245, 245));
        salasListModel = new DefaultListModel<>();

        JPanel menuPanel = criarMenuPanel();
        mainPanel.add(menuPanel, "menu");

        JPanel cadastroSalaPanel = criarCadastroSalaPanel();
        mainPanel.add(cadastroSalaPanel, "cadastroSala");

        JPanel listarSalasPanel = criarListarSalasPanel();
        mainPanel.add(listarSalasPanel, "listarSalas");

        JPanel cadastroAlunoPanel = criarCadastroAlunoPanel();
        mainPanel.add(cadastroAlunoPanel, "cadastroAluno");

        frame.add(mainPanel);
        frame.setVisible(true);
    }

    private static JPanel criarMenuPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new EmptyBorder(20, 20, 20, 20));
        panel.setBackground(new Color(245, 245, 245));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridwidth = GridBagConstraints.REMAINDER;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(15, 0, 15, 0);

        JLabel titulo = new JLabel("Sistema Escolar");
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 32));
        titulo.setForeground(new Color(0, 90, 158));
        panel.add(titulo, gbc);

        JButton cadastrarSalaBtn = criarBotaoEstilizado("Cadastrar Nova Sala", new Color(0, 105, 92));
        cadastrarSalaBtn.setPreferredSize(new Dimension(300, 50));
        cadastrarSalaBtn.addActionListener(e -> cardLayout.show(mainPanel, "cadastroSala"));
        panel.add(cadastrarSalaBtn, gbc);

        JButton verSalasBtn = criarBotaoEstilizado("Ver Salas Cadastradas", new Color(0, 90, 158));
        verSalasBtn.setPreferredSize(new Dimension(300, 50));
        verSalasBtn.addActionListener(e -> {
            atualizarListaSalas();
            cardLayout.show(mainPanel, "listarSalas");
        });
        panel.add(verSalasBtn, gbc);

        return panel;
    }

    private static JPanel criarCadastroSalaPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new EmptyBorder(20, 20, 20, 20));
        panel.setBackground(new Color(245, 245, 245));

        JLabel titulo = new JLabel("Cadastrar Nova Sala");
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 24));
        titulo.setForeground(new Color(0, 90, 158));
        titulo.setHorizontalAlignment(SwingConstants.CENTER);
        panel.add(titulo, BorderLayout.NORTH);

        JPanel formPanel = new JPanel(new GridBagLayout());
        formPanel.setBorder(new EmptyBorder(30, 150, 30, 150));
        formPanel.setBackground(new Color(245, 245, 245));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.LINE_START;
        gbc.insets = new Insets(10, 5, 10, 15);

        JLabel anoLabel = new JLabel("Ano da Sala:");
        anoLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        formPanel.add(anoLabel, gbc);

        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        JTextField anoField = criarCampoTextoEstilizado();
        formPanel.add(anoField, gbc);

        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.fill = GridBagConstraints.NONE;
        gbc.weightx = 0;
        JLabel identificacaoLabel = new JLabel("Identificação (ex: A, B, C):");
        identificacaoLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        formPanel.add(identificacaoLabel, gbc);

        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        JTextField identificacaoField = criarCampoTextoEstilizado();
        formPanel.add(identificacaoField, gbc);

        panel.add(formPanel, BorderLayout.CENTER);

        JPanel botoesPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 20, 20));
        botoesPanel.setBackground(new Color(245, 245, 245));
        
        JButton cadastrarBtn = criarBotaoEstilizado("Cadastrar Sala", new Color(0, 105, 92));
        cadastrarBtn.setPreferredSize(new Dimension(200, 45));
        cadastrarBtn.addActionListener(e -> {
            try {
                int ano = Integer.parseInt(anoField.getText());
                String identificacao = identificacaoField.getText().trim().toUpperCase();
                
                if (identificacao.isEmpty()) {
                    JOptionPane.showMessageDialog(frame, "Informe a identificação da sala!");
                    return;
                }
                
                String nomeSala = ano + "º Ano " + identificacao;
                cadastrarSala(nomeSala);
                
                JOptionPane.showMessageDialog(frame, "Sala cadastrada com sucesso!");
                anoField.setText("");
                identificacaoField.setText("");
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(frame, "Ano deve ser um número válido!");
            } catch (SQLException ex) {
                JOptionPane.showMessageDialog(frame, "Erro ao cadastrar sala: " + ex.getMessage());
            }
        });
        botoesPanel.add(cadastrarBtn);
        
        JButton voltarBtn = criarBotaoEstilizado("Voltar ao Menu", new Color(158, 0, 0));
        voltarBtn.setPreferredSize(new Dimension(200, 45));
        voltarBtn.addActionListener(e -> cardLayout.show(mainPanel, "menu"));
        botoesPanel.add(voltarBtn);
        
        panel.add(botoesPanel, BorderLayout.SOUTH);

        return panel;
    }

    private static JPanel criarListarSalasPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));
        panel.setBackground(new Color(245, 245, 245));

        JLabel titulo = new JLabel("Salas Cadastradas");
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 24));
        titulo.setForeground(new Color(0, 90, 158));
        titulo.setHorizontalAlignment(SwingConstants.CENTER);
        panel.add(titulo, BorderLayout.NORTH);

        JList<String> salasList = new JList<>(salasListModel);
        salasList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        salasList.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        salasList.setFixedCellHeight(45);
        salasList.setBackground(Color.WHITE);
        salasList.setSelectionBackground(new Color(200, 220, 255));
        
        JScrollPane scrollPane = new JScrollPane(salasList);
        scrollPane.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createEmptyBorder(20, 50, 20, 50),
            BorderFactory.createLineBorder(new Color(200, 200, 200), 2)));
        panel.add(scrollPane, BorderLayout.CENTER);

        // Declaração do botoesPanel antes de usá-lo
        JPanel botoesPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 15));
        botoesPanel.setBackground(new Color(245, 245, 245));
        
        JButton adicionarAlunoBtn = criarBotaoEstilizado("Adicionar Aluno", new Color(0, 105, 92));
        adicionarAlunoBtn.setPreferredSize(new Dimension(180, 40));
        adicionarAlunoBtn.addActionListener(e -> {
            int selectedIndex = salasList.getSelectedIndex();
            if (selectedIndex == -1) {
                JOptionPane.showMessageDialog(frame, "Selecione uma sala primeiro!");
                return;
            }
            
            String salaSelecionada = salasListModel.getElementAt(selectedIndex);
            String[] parts = salaSelecionada.split(" - ");
            int salaId = Integer.parseInt(parts[0]);
            
            abrirCadastroAluno(salaId);
        });
        botoesPanel.add(adicionarAlunoBtn);
        
        JButton verAlunosBtn = criarBotaoEstilizado("Ver Alunos", new Color(0, 90, 158));
        verAlunosBtn.setPreferredSize(new Dimension(180, 40));
        verAlunosBtn.addActionListener(e -> {
            int selectedIndex = salasList.getSelectedIndex();
            if (selectedIndex == -1) {
                JOptionPane.showMessageDialog(frame, "Selecione uma sala primeiro!");
                return;
            }
            
            String salaSelecionada = salasListModel.getElementAt(selectedIndex);
            String[] parts = salaSelecionada.split(" - ");
            int salaId = Integer.parseInt(parts[0]);
            String nomeSala = parts[1].split(" \\(")[0];
            
            mostrarAlunos(salaId, nomeSala);
        });
        botoesPanel.add(verAlunosBtn);
        
        JButton gerarGraficosBtn = criarBotaoEstilizado("Gerar Gráficos", new Color(128, 0, 128));
        gerarGraficosBtn.setPreferredSize(new Dimension(180, 40));
        gerarGraficosBtn.addActionListener(e -> {
            int selectedIndex = salasList.getSelectedIndex();
            if (selectedIndex == -1) {
                JOptionPane.showMessageDialog(frame, "Selecione uma sala primeiro!");
                return;
            }
            
            String salaSelecionada = salasListModel.getElementAt(selectedIndex);
            String[] parts = salaSelecionada.split(" - ");
            int salaId = Integer.parseInt(parts[0]);
            String nomeSala = parts[1].split(" \\(")[0];
            
            gerarGraficos(salaId, nomeSala);
        });
        botoesPanel.add(gerarGraficosBtn);
        
        // Adicione o botão de importar do Excel
        JButton importarExcelBtn = criarBotaoEstilizado("Importar do Excel", new Color(0, 77, 64));
        importarExcelBtn.setPreferredSize(new Dimension(180, 40));
        importarExcelBtn.addActionListener(e -> {
            int selectedIndex = salasList.getSelectedIndex();
            if (selectedIndex == -1) {
                JOptionPane.showMessageDialog(frame, "Selecione uma sala primeiro!");
                return;
            }
            
            String salaSelecionada = salasListModel.getElementAt(selectedIndex);
            String[] parts = salaSelecionada.split(" - ");
            int salaId = Integer.parseInt(parts[0]);
            String nomeSala = parts[1].split(" \\(")[0];
            
            importarDeExcel(salaId, nomeSala);
        });
        botoesPanel.add(importarExcelBtn);
        
        JButton voltarBtn = criarBotaoEstilizado("Voltar ao Menu", new Color(158, 0, 0));
        voltarBtn.setPreferredSize(new Dimension(180, 40));
        voltarBtn.addActionListener(e -> cardLayout.show(mainPanel, "menu"));
        botoesPanel.add(voltarBtn);
        
        panel.add(botoesPanel, BorderLayout.SOUTH);

        return panel;
    }

    @SuppressWarnings("null")
	private static JPanel criarCadastroAlunoPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));
        panel.setBackground(new Color(245, 245, 245));

        JLabel titulo = new JLabel("Cadastrar Novo Aluno");
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 24));
        titulo.setForeground(new Color(0, 90, 158));
        titulo.setHorizontalAlignment(SwingConstants.CENTER);
        panel.add(titulo, BorderLayout.NORTH);

        JPanel formPanel = new JPanel(new GridBagLayout());
        formPanel.setBorder(new EmptyBorder(30, 150, 30, 150));
        formPanel.setBackground(new Color(245, 245, 245));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.LINE_START;
        gbc.insets = new Insets(10, 5, 10, 15);

        JLabel nomeLabel = new JLabel("Nome do Aluno:");
        nomeLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        formPanel.add(nomeLabel, gbc);

        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        JTextField nomeField = criarCampoTextoEstilizado();
        formPanel.add(nomeField, gbc);

        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.fill = GridBagConstraints.NONE;
        gbc.weightx = 0;
        JLabel fezProvaLabel = new JLabel("Fez a prova?");
        fezProvaLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        formPanel.add(fezProvaLabel, gbc);

        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        JCheckBox fezProvaCheck = new JCheckBox("Sim");
        fezProvaCheck.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        fezProvaCheck.setBackground(new Color(245, 245, 245));
        fezProvaCheck.addActionListener(e -> {
            if (!fezProvaCheck.isSelected()) {
                JLabel notaField = null;
				notaField.setText("");
            }
        });
        formPanel.add(fezProvaCheck, gbc);

        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.fill = GridBagConstraints.NONE;
        gbc.weightx = 0;
        JLabel notaLabel = new JLabel("Nota:");
        notaLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        formPanel.add(notaLabel, gbc);

        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        JTextField notaField = criarCampoTextoEstilizado();
        formPanel.add(notaField, gbc);

        panel.add(formPanel, BorderLayout.CENTER);

        JPanel botoesPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 20, 20));
        botoesPanel.setBackground(new Color(245, 245, 245));
        
        JButton cadastrarBtn = criarBotaoEstilizado("Cadastrar Aluno", new Color(0, 105, 92));
        cadastrarBtn.setPreferredSize(new Dimension(200, 45));
        cadastrarBtn.addActionListener(e -> {
            try {
                String nome = nomeField.getText().trim();
                boolean fezProva = fezProvaCheck.isSelected();
                float nota = 0;
                
                if (nome.isEmpty()) {
                    JOptionPane.showMessageDialog(frame, "Informe o nome do aluno!");
                    return;
                }
                
                if (fezProva) {
                    try {
                        nota = Float.parseFloat(notaField.getText());
                        if (nota < 0 || nota > 10) {
                            JOptionPane.showMessageDialog(frame, "Nota deve estar entre 0 e 10!");
                            return;
                        }
                    } catch (NumberFormatException ex) {
                        JOptionPane.showMessageDialog(frame, "Nota deve ser um número válido!");
                        return;
                    }
                }
                
                int salaId = (Integer) panel.getClientProperty("salaId");
                cadastrarAluno(salaId, nome, fezProva, nota);
                
                JOptionPane.showMessageDialog(frame, "Aluno cadastrado com sucesso!");
                nomeField.setText("");
                notaField.setText("");
                fezProvaCheck.setSelected(false);
            } catch (SQLException ex) {
                JOptionPane.showMessageDialog(frame, "Erro ao cadastrar aluno: " + ex.getMessage());
            }
        });
        botoesPanel.add(cadastrarBtn);
        
        JButton voltarBtn = criarBotaoEstilizado("Voltar", new Color(158, 0, 0));
        voltarBtn.setPreferredSize(new Dimension(200, 45));
        voltarBtn.addActionListener(e -> {
            atualizarListaSalas();
            cardLayout.show(mainPanel, "listarSalas");
        });
        botoesPanel.add(voltarBtn);
        
        panel.add(botoesPanel, BorderLayout.SOUTH);

        return panel;
    }

    private static JButton criarBotaoEstilizado(String texto, Color corFundo) {
        JButton botao = new JButton(texto);
        botao.setFont(new Font("Segoe UI", Font.BOLD, 14));
        botao.setBackground(corFundo);
        botao.setForeground(Color.WHITE);
        botao.setFocusPainted(false);
        botao.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));
        botao.setCursor(new Cursor(Cursor.HAND_CURSOR));
        return botao;
    }

    private static JTextField criarCampoTextoEstilizado() {
        JTextField campo = new JTextField();
        campo.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        campo.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(200, 200, 200), 1),
            BorderFactory.createEmptyBorder(8, 12, 8, 12)
        ));
        campo.setBackground(Color.WHITE);
        return campo;
    }

    private static void criarTabelas() throws SQLException {
        Statement stmt = connection.createStatement();
        stmt.execute("CREATE TABLE IF NOT EXISTS salas (" +
                     "id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                     "nome TEXT UNIQUE)");
        
        stmt.execute("CREATE TABLE IF NOT EXISTS alunos (" +
                     "id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                     "nome TEXT, " +
                     "fez_prova BOOLEAN, " +
                     "nota REAL, " +
                     "sala_id INTEGER, " +
                     "FOREIGN KEY(sala_id) REFERENCES salas(id))");
        stmt.close();
    }

    private static void cadastrarSala(String nomeSala) throws SQLException {
        PreparedStatement ps = connection.prepareStatement(
            "INSERT INTO salas (nome) VALUES (?)");
        ps.setString(1, nomeSala);
        ps.executeUpdate();
        ps.close();
    }

    private static void cadastrarAluno(int salaId, String nome, boolean fezProva, float nota) throws SQLException {
        PreparedStatement ps = connection.prepareStatement(
            "INSERT INTO alunos (nome, fez_prova, nota, sala_id) VALUES (?, ?, ?, ?)");
        ps.setString(1, nome);
        ps.setBoolean(2, fezProva);
        ps.setFloat(3, nota);
        ps.setInt(4, salaId);
        ps.executeUpdate();
        ps.close();
    }

    private static void excluirAluno(int alunoId) throws SQLException {
        PreparedStatement ps = connection.prepareStatement(
            "DELETE FROM alunos WHERE id = ?");
        ps.setInt(1, alunoId);
        ps.executeUpdate();
        ps.close();
    }

    private static void atualizarListaSalas() {
        salasListModel.clear();
        
        try {
            Statement stmt = connection.createStatement();
            ResultSet rs = stmt.executeQuery("SELECT id, nome FROM salas ORDER BY nome");
            
            while (rs.next()) {
                int id = rs.getInt("id");
                String nome = rs.getString("nome");
                
                PreparedStatement ps = connection.prepareStatement(
                    "SELECT COUNT(*) as total FROM alunos WHERE sala_id = ?");
                ps.setInt(1, id);
                ResultSet rsCount = ps.executeQuery();
                int totalAlunos = rsCount.getInt("total");
                
                salasListModel.addElement(id + " - " + nome + " (" + totalAlunos + " alunos)");
                ps.close();
            }
            
            stmt.close();
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(frame, "Erro ao carregar salas: " + ex.getMessage());
        }
    }

    private static void abrirCadastroAluno(int salaId) {
        JPanel cadastroAlunoPanel = (JPanel) mainPanel.getComponent(3);
        cadastroAlunoPanel.putClientProperty("salaId", salaId);
        cardLayout.show(mainPanel, "cadastroAluno");
    }

    private static void mostrarAlunos(int salaId, String nomeSala) {
        JDialog dialog = new JDialog(frame, "Alunos da Sala " + nomeSala, true);
        dialog.setSize(700, 550);
        dialog.setLocationRelativeTo(frame);
        
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));
        panel.setBackground(new Color(245, 245, 245));
        
        JPanel alunosPanel = new JPanel(new BorderLayout());
        alunosPanel.setBackground(Color.WHITE);
        alunosPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        
        DefaultListModel<String> alunosListModel = new DefaultListModel<>();
        JList<String> alunosList = new JList<>(alunosListModel);
        alunosList.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        alunosList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        alunosList.setFixedCellHeight(40);
        alunosList.setBackground(Color.WHITE);
        alunosList.setSelectionBackground(new Color(200, 220, 255));
        
        Runnable carregarAlunos = () -> {
            alunosListModel.clear(); 
            
            try {
                PreparedStatement ps = connection.prepareStatement(
                    "SELECT id, nome, fez_prova, nota FROM alunos WHERE sala_id = ? ORDER BY nome");
                ps.setInt(1, salaId);
                ResultSet rs = ps.executeQuery();
                
                while (rs.next()) {
                    int id = rs.getInt("id");
                    String nome = rs.getString("nome");
                    boolean fezProva = rs.getBoolean("fez_prova");
                    float nota = rs.getFloat("nota");
                    
                    String alunoInfo = nome;
                    if (fezProva) {
                        alunoInfo += " - Nota: " + String.format("%.1f", nota);
                    } else {
                        alunoInfo += " - Não fez a prova";
                    }
                    
                    alunosListModel.addElement(alunoInfo);
                }
                
                if (alunosListModel.isEmpty()) {
                    alunosListModel.addElement("Nenhum aluno cadastrado nesta sala.");
                }
                
                ps.close();
            } catch (SQLException ex) {
                alunosListModel.addElement("Erro ao carregar alunos: " + ex.getMessage());
            }
        };
        
        carregarAlunos.run();
        
        JScrollPane scrollPane = new JScrollPane(alunosList);
        scrollPane.setBorder(BorderFactory.createEmptyBorder());
        alunosPanel.add(scrollPane, BorderLayout.CENTER);
        
        panel.add(alunosPanel, BorderLayout.CENTER);
        
        JPanel botoesPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 15));
        botoesPanel.setBackground(new Color(245, 245, 245));
        
        JButton excluirBtn = criarBotaoEstilizado("Excluir Aluno", new Color(158, 0, 0));
        excluirBtn.setPreferredSize(new Dimension(180, 40));
        excluirBtn.addActionListener(e -> {
            int selectedIndex = alunosList.getSelectedIndex();
            if (selectedIndex == -1) {
                JOptionPane.showMessageDialog(dialog, "Selecione um aluno primeiro!");
                return;
            }
            
            try {
                PreparedStatement ps = connection.prepareStatement(
                    "SELECT id FROM alunos WHERE sala_id = ? ORDER BY nome");
                ps.setInt(1, salaId);
                ResultSet rs = ps.executeQuery();
                
                int alunoId = -1;
                for (int i = 0; i <= selectedIndex; i++) {
                    if (!rs.next()) break;
                    alunoId = rs.getInt("id");
                }
                
                if (alunoId != -1) {
                    int confirm = JOptionPane.showConfirmDialog(
                        dialog, 
                        "Tem certeza que deseja excluir este aluno?", 
                        "Confirmar Exclusão", 
                        JOptionPane.YES_NO_OPTION);
                    
                    if (confirm == JOptionPane.YES_OPTION) {
                        excluirAluno(alunoId);
                        
                        carregarAlunos.run();
                        atualizarListaSalas();
                    }
                }
                
                ps.close();
            } catch (SQLException ex) {
                JOptionPane.showMessageDialog(dialog, "Erro ao excluir aluno: " + ex.getMessage());
            }
        });
        botoesPanel.add(excluirBtn);
        
        JButton fecharBtn = criarBotaoEstilizado("Fechar", new Color(100, 100, 100));
        fecharBtn.setPreferredSize(new Dimension(180, 40));
        fecharBtn.addActionListener(e -> dialog.dispose());
        botoesPanel.add(fecharBtn);
        
        panel.add(botoesPanel, BorderLayout.SOUTH);
        
        dialog.add(panel);
        dialog.setVisible(true);
    }

    private static void gerarGraficos(int salaId, String nomeSala) {
        File csvFile = null;
        File outputImage = null;
        
        try {
            // 1. Criar arquivo CSV temporário
            csvFile = File.createTempFile("alunos_", ".csv");
            csvFile.deleteOnExit();
            
            // 2. Escrever dados no CSV com tratamento completo
            try (PrintWriter writer = new PrintWriter(csvFile, "UTF-8")) {
                writer.println("nome,fez_prova,nota");  // Cabeçalho
                
                PreparedStatement ps = connection.prepareStatement(
                    "SELECT nome, fez_prova, nota FROM alunos WHERE sala_id = ? ORDER BY nome");
                ps.setInt(1, salaId);
                ResultSet rs = ps.executeQuery();
                
                while (rs.next()) {
                    String nome = rs.getString("nome");
                    boolean fezProva = rs.getBoolean("fez_prova");
                    float nota = rs.getFloat("nota");
                    
                    // Tratamento completo dos dados:
                    // - Escapar aspas e quebras de linha no nome
                    String nomeEscapado = nome.replace("\"", "\"\"")
                                            .replace("\n", " ")
                                            .replace("\r", " ");
                    
                    // - Converter boolean para 1/0
                    int provaInt = fezProva ? 1 : 0;
                    
                    // - Garantir formato decimal correto (usando ponto)
                    String notaFormatada = String.format(Locale.US, "%.1f", nota);
                    
                    // Escrever linha formatada corretamente
                    writer.printf("\"%s\",%d,%s%n", nomeEscapado, provaInt, notaFormatada);
                }
                rs.close();
                ps.close();
            } catch (SQLException e) {
                JOptionPane.showMessageDialog(frame, 
                    "Erro ao acessar dados dos alunos: " + e.getMessage(),
                    "Erro no Banco de Dados",
                    JOptionPane.ERROR_MESSAGE);
                return;
            }
            
            // 3. Criar arquivo de saída para a imagem
            outputImage = File.createTempFile("graficos_", ".png");
            outputImage.deleteOnExit();
            
            // 4. Executar script Python
            ProcessBuilder pb = new ProcessBuilder(
                "python",
                "gerar_graficos.py",
                csvFile.getAbsolutePath(),
                nomeSala.replace("\"", "\\\""),  // Escapar aspas no nome da sala
                outputImage.getAbsolutePath()
            );
            
            // Configurar ambiente de execução
            pb.redirectErrorStream(true);
            pb.directory(new File(System.getProperty("user.dir")));
            
            // Executar e capturar saída
            Process process = pb.start();
            
            // Ler saída para debug
            StringBuilder output = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(
                    new InputStreamReader(process.getInputStream(), "UTF-8"))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    output.append(line).append("\n");
                }
            }
            
            // Verificar resultado
            int exitCode = process.waitFor();
            
            if (exitCode != 0) {
                throw new RuntimeException("Script Python retornou código de erro: " + exitCode + 
                                        "\nSaída:\n" + output.toString());
            }
            
            // 5. Mostrar resultados
            JDialog dialog = new JDialog(frame, "Gráficos - " + nomeSala, true);
            dialog.setSize(900, 700);
            dialog.setLocationRelativeTo(frame);
            
            // Carregar imagem com verificação
            ImageIcon icon = new ImageIcon(outputImage.getAbsolutePath());
            if (icon.getImageLoadStatus() != MediaTracker.COMPLETE) {
                throw new IOException("Falha ao carregar imagem gerada");
            }
            
            JLabel label = new JLabel(icon);
            JScrollPane scrollPane = new JScrollPane(label);
            scrollPane.setPreferredSize(new Dimension(850, 600));
            
            JButton fecharBtn = criarBotaoEstilizado("Fechar", new Color(100, 100, 100));
            fecharBtn.addActionListener(e -> dialog.dispose());
            
            JPanel panel = new JPanel(new BorderLayout());
            panel.add(scrollPane, BorderLayout.CENTER);
            
            JPanel btnPanel = new JPanel();
            btnPanel.add(fecharBtn);
            panel.add(btnPanel, BorderLayout.SOUTH);
            
            dialog.add(panel);
            dialog.setVisible(true);
            
        } catch (IOException e) {
            JOptionPane.showMessageDialog(frame,
                "Erro de I/O ao gerar gráficos: " + e.getMessage(),
                "Erro de Sistema",
                JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            JOptionPane.showMessageDialog(frame,
                "Operação interrompida ao gerar gráficos",
                "Interrupção",
                JOptionPane.WARNING_MESSAGE);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(frame,
                "Erro inesperado ao gerar gráficos:\n" + e.getMessage(),
                "Erro",
                JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        } finally {
            // Limpeza adicional se necessário
            if (csvFile != null && csvFile.exists()) {
                //csvFile.delete(); // Manter temporário para debug
            }
        }
    }
    
    private static void importarDeExcel(int salaId, String nomeSala) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Selecione o arquivo Excel");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Arquivos Excel", "xlsx", "xls"));
        
        int returnValue = fileChooser.showOpenDialog(frame);
        if (returnValue != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File selectedFile = fileChooser.getSelectedFile();
        
        // Diálogo de progresso com botão de cancelamento
        JDialog progressDialog = new JDialog(frame, "Importando alunos...", true);
        progressDialog.setSize(400, 150);
        progressDialog.setLocationRelativeTo(frame);
        progressDialog.setLayout(new BorderLayout());
        
        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);
        progressBar.setString("Processando arquivo...");
        progressBar.setStringPainted(true);
        
        JButton cancelButton = new JButton("Cancelar");
        AtomicBoolean cancelRequested = new AtomicBoolean(false);
        
        cancelButton.addActionListener(e -> {
            cancelRequested.set(true);
            progressDialog.dispose();
        });
        
        JPanel buttonPanel = new JPanel();
        buttonPanel.add(cancelButton);
        
        progressDialog.add(progressBar, BorderLayout.CENTER);
        progressDialog.add(buttonPanel, BorderLayout.SOUTH);
        
        // Thread de processamento
        new SwingWorker<Void, Integer>() {
            private int alunosImportados = 0;
            private Exception error = null;
            
            @Override
            protected Void doInBackground() throws Exception {
                SwingUtilities.invokeLater(() -> progressDialog.setVisible(true));
                
                try (Workbook workbook = WorkbookFactory.create(selectedFile)) {
                    Sheet sheet = workbook.getSheetAt(0);
                    DataFormatter formatter = new DataFormatter();
                    
                    // Verifica estrutura básica
                    if (sheet.getPhysicalNumberOfRows() <= 1) {
                        throw new Exception("O arquivo não contém dados além do cabeçalho");
                    }
                    
                    // Processa cada linha
                    for (Row row : sheet) {
                        if (cancelRequested.get()) {
                            break;
                        }
                        
                        if (row.getRowNum() == 0) continue; // Pula cabeçalho
                        
                        String nome = formatter.formatCellValue(row.getCell(0)).trim();
                        if (nome.isEmpty()) continue;
                        
                        String notaStr = formatter.formatCellValue(row.getCell(1)).trim();
                        boolean fezProva = !notaStr.isEmpty();
                        float nota = 0;
                        
                        try {
                            if (fezProva) {
                                nota = Float.parseFloat(notaStr);
                                if (nota < 0 || nota > 10) {
                                    System.out.println("Nota inválida para " + nome + ": " + nota);
                                    fezProva = false;
                                }
                            }
                            
                            cadastrarAluno(salaId, nome, fezProva, nota);
                            alunosImportados++;
                            
                            // Atualiza progresso a cada 10 registros
                            if (alunosImportados % 10 == 0) {
                                publish(alunosImportados);
                            }
                            
                        } catch (SQLException e) {
                            System.err.println("Erro ao cadastrar aluno " + nome + ": " + e.getMessage());
                        }
                    }
                    
                } catch (Exception e) {
                    this.error = e;
                }
                
                return null;
            }
            
            @Override
            protected void process(List<Integer> chunks) {
                int lastProgress = chunks.get(chunks.size() - 1);
                progressBar.setString(lastProgress + " alunos importados...");
            }
            
            @Override
            protected void done() {
                progressDialog.dispose();
                
                if (cancelRequested.get()) {
                    JOptionPane.showMessageDialog(frame, 
                        "Importação cancelada! " + alunosImportados + " alunos foram importados.",
                        "Importação interrompida",
                        JOptionPane.INFORMATION_MESSAGE);
                    return;
                }
                
                if (error != null) {
                    JOptionPane.showMessageDialog(frame, 
                        "Erro durante importação: " + error.getMessage(),
                        "Erro",
                        JOptionPane.ERROR_MESSAGE);
                    return;
                }
                
                JOptionPane.showMessageDialog(frame, 
                    alunosImportados + " alunos importados com sucesso para a sala " + nomeSala + "!",
                    "Importação concluída",
                    JOptionPane.INFORMATION_MESSAGE);
                
                atualizarListaSalas();
            }
        }.execute();
    }

    private static int processarArquivoExcel(File excelFile, int salaId) throws Exception {
        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(excelFile)) {
            workbook = WorkbookFactory.create(fis);
        }
        
        Sheet sheet = workbook.getSheetAt(0); // Pega a primeira planilha
        int alunosImportados = 0;
        
        // Encontrar índices das colunas
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            throw new Exception("O arquivo Excel não contém cabeçalhos na primeira linha");
        }
        
        int nomeCol = -1;
        int notaCol = -1;
        
        for (Cell cell : headerRow) {
            String header = cell.getStringCellValue().trim().toLowerCase();
            if (header.equals("nome")) {
                nomeCol = cell.getColumnIndex();
            } else if (header.equals("nota")) {
                notaCol = cell.getColumnIndex();
            }
        }
        
        if (nomeCol == -1) {
            throw new Exception("Coluna 'nome' não encontrada no arquivo Excel");
        }
        
        if (notaCol == -1) {
            throw new Exception("Coluna 'nota' não encontrada no arquivo Excel");
        }
        
        // Processar linhas
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            
            Cell nomeCell = row.getCell(nomeCol);
            Cell notaCell = row.getCell(notaCol);
            
            if (nomeCell == null || nomeCell.getCellType() == CellType.BLANK) {
                continue; // Pula linhas sem nome
            }
            
            String nome = nomeCell.getStringCellValue().trim();
            if (nome.isEmpty()) continue;
            
            boolean fezProva = false;
            double nota = 0;
            
            if (notaCell != null && notaCell.getCellType() != CellType.BLANK) {
                try {
                    fezProva = true;
                    nota = notaCell.getNumericCellValue();
                    
                    // Validar nota
                    if (nota < 0 || nota > 10) {
                        System.err.println("Nota inválida para " + nome + ": " + nota + " - será considerada como não fez a prova");
                        fezProva = false;
                        nota = 0;
                    }
                } catch (Exception e) {
                    System.err.println("Erro ao ler nota para " + nome + " - será considerada como não fez a prova");
                    fezProva = false;
                    nota = 0;
                }
            }
            
            // Inserir no banco de dados
            try {
                cadastrarAluno(salaId, nome, fezProva, (float) nota);
                alunosImportados++;
            } catch (SQLException e) {
                System.err.println("Erro ao cadastrar aluno " + nome + ": " + e.getMessage());
            }
        }
        
        workbook.close();
        return alunosImportados;
    }
}