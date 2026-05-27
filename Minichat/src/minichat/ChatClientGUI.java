package minichat;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.ConnectException;
import java.net.Socket;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;

public class ChatClientGUI extends JFrame {
    private static final String SERVER_HOST = "localhost";
    private static final int SERVER_PORT = 9999;

    private final JTextArea chatArea = new JTextArea();
    private final JTextField usernameField = new JTextField();
    private final JTextField messageField = new JTextField();
    private final JButton connectButton = new JButton("Connect");
    private final JButton sendButton = new JButton("Send");
    private final JLabel statusLabel = new JLabel("Disconnected");

    private Socket socket;
    private BufferedReader serverReader;
    private PrintWriter serverWriter;
    private Thread receiveThread;
    private boolean connected;

    public ChatClientGUI() {
        setTitle("MiniChat");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(560, 420));
        setLocationRelativeTo(null);

        buildLayout();
        bindActions();
        setDisconnectedState();
    }

    private void buildLayout() {
        JPanel rootPanel = new JPanel(new BorderLayout(12, 12));
        rootPanel.setBorder(BorderFactory.createEmptyBorder(14, 14, 14, 14));
        rootPanel.setBackground(new Color(245, 247, 250));

        JPanel topPanel = new JPanel(new BorderLayout(8, 8));
        topPanel.setOpaque(false);
        JLabel usernameLabel = new JLabel("Username");
        usernameField.setFont(new Font("SansSerif", Font.PLAIN, 14));
        usernameField.setToolTipText("Enter your chat name");
        connectButton.setToolTipText("Connect to MiniChat Server");
        topPanel.add(usernameLabel, BorderLayout.WEST);
        topPanel.add(usernameField, BorderLayout.CENTER);
        topPanel.add(connectButton, BorderLayout.EAST);

        chatArea.setEditable(false);
        chatArea.setLineWrap(true);
        chatArea.setWrapStyleWord(true);
        chatArea.setFont(new Font("Monospaced", Font.PLAIN, 14));
        chatArea.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        JScrollPane chatScrollPane = new JScrollPane(chatArea);
        chatScrollPane.setBorder(BorderFactory.createLineBorder(new Color(210, 215, 222)));

        JPanel bottomPanel = new JPanel(new BorderLayout(8, 8));
        bottomPanel.setOpaque(false);
        messageField.setFont(new Font("SansSerif", Font.PLAIN, 14));
        messageField.setToolTipText("Type a message");
        sendButton.setToolTipText("Send message");
        bottomPanel.add(statusLabel, BorderLayout.NORTH);
        bottomPanel.add(messageField, BorderLayout.CENTER);
        bottomPanel.add(sendButton, BorderLayout.EAST);

        rootPanel.add(topPanel, BorderLayout.NORTH);
        rootPanel.add(chatScrollPane, BorderLayout.CENTER);
        rootPanel.add(bottomPanel, BorderLayout.SOUTH);
        setContentPane(rootPanel);
    }

    private void bindActions() {
        connectButton.addActionListener(event -> {
            if (connected) {
                disconnect();
            } else {
                connect();
            }
        });

        sendButton.addActionListener(event -> sendMessage());
        messageField.addActionListener(event -> sendMessage());

        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent event) {
                disconnect();
            }
        });
    }

    private void connect() {
        String username = usernameField.getText().trim();

        if (username.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please enter your username.");
            usernameField.requestFocus();
            return;
        }

        try {
            socket = new Socket(SERVER_HOST, SERVER_PORT);
            serverReader = new BufferedReader(new InputStreamReader(socket.getInputStream()));
            serverWriter = new PrintWriter(socket.getOutputStream(), true);

            connected = true;
            setConnectedState();

            receiveThread = new Thread(this::receiveMessages);
            receiveThread.start();

            serverWriter.println(username);
        } catch (IOException e) {
            connected = false;
            closeResources();

            if (e instanceof ConnectException) {
                showError("Cannot connect to MiniChat Server. Please run ChatServer first.");
            } else {
                showError("Client error: " + e.getMessage());
            }
        }
    }

    private void receiveMessages() {
        try {
            String serverMessage;

            while ((serverMessage = serverReader.readLine()) != null) {
                if (serverMessage.equals("Enter your username:")) {
                    continue;
                }

                appendMessage(serverMessage);
            }
        } catch (IOException e) {
            if (connected) {
                appendMessage("Disconnected from server.");
            }
        } finally {
            SwingUtilities.invokeLater(this::setDisconnectedState);
            closeResources();
        }
    }

    private void sendMessage() {
        String message = messageField.getText().trim();

        if (!connected || message.isEmpty()) {
            return;
        }

        serverWriter.println(message);
        appendMessage("Me: " + message);
        messageField.setText("");

        if (message.equalsIgnoreCase("/quit")) {
            connected = false;
            closeResources();
            setDisconnectedState();
        }
    }

    private void disconnect() {
        if (connected && serverWriter != null) {
            serverWriter.println("/quit");
        }

        connected = false;
        closeResources();
        setDisconnectedState();
    }

    private void closeResources() {
        try {
            if (serverReader != null) {
                serverReader.close();
            }

            if (serverWriter != null) {
                serverWriter.close();
            }

            if (socket != null && !socket.isClosed()) {
                socket.close();
            }
        } catch (IOException e) {
            appendMessage("Close connection error: " + e.getMessage());
        }
    }

    private void setConnectedState() {
        statusLabel.setText("Connected to " + SERVER_HOST + ":" + SERVER_PORT);
        usernameField.setEnabled(false);
        connectButton.setText("Disconnect");
        messageField.setEnabled(true);
        sendButton.setEnabled(true);
        messageField.requestFocus();
    }

    private void setDisconnectedState() {
        connected = false;
        statusLabel.setText("Disconnected");
        usernameField.setEnabled(true);
        connectButton.setText("Connect");
        messageField.setEnabled(false);
        sendButton.setEnabled(false);
    }

    private void appendMessage(String message) {
        SwingUtilities.invokeLater(() -> {
            chatArea.append(message + System.lineSeparator());
            chatArea.setCaretPosition(chatArea.getDocument().getLength());
        });
    }

    private void showError(String message) {
        JOptionPane.showMessageDialog(this, message, "MiniChat", JOptionPane.ERROR_MESSAGE);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ChatClientGUI clientGUI = new ChatClientGUI();
            clientGUI.setVisible(true);
        });
    }
}
