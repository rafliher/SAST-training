<?php
// Vulnerable test application for LLM-SAST demo
session_start();

// Database connection
$db = new mysqli("localhost", "root", "", "app_db");

// LOGIN - SQL Injection vulnerable
function login($db) {
    $username = $_POST['username'];
    $password = $_POST['password'];

    $query = "SELECT * FROM users WHERE username = '$username' AND password = '$password'";
    $result = $db->query($query);

    if ($result->num_rows > 0) {
        $user = $result->fetch_assoc();
        $_SESSION['user_id'] = $user['id'];
        $_SESSION['role'] = $user['role'];
        header("Location: dashboard.php");
    } else {
        echo "Invalid credentials";
    }
}

// PROFILE - XSS vulnerable
function show_profile($db) {
    $user_id = $_GET['id'];
    $query = "SELECT * FROM users WHERE id = $user_id";
    $result = $db->query($query);
    $user = $result->fetch_assoc();

    echo "<h1>Welcome, " . $user['name'] . "</h1>";
    echo "<p>Bio: " . $user['bio'] . "</p>";
    echo "<p>Email: " . $user['email'] . "</p>";
}

// FILE UPLOAD - Unrestricted upload
function upload_avatar() {
    $target_dir = "uploads/";
    $target_file = $target_dir . basename($_FILES["avatar"]["name"]);
    move_uploaded_file($_FILES["avatar"]["tmp_name"], $target_file);
    echo "File uploaded: " . $target_file;
}

// ADMIN PANEL - Missing authorization check (IDOR)
function delete_user($db) {
    $user_id = $_GET['delete_id'];
    $db->query("DELETE FROM users WHERE id = $user_id");
    echo "User deleted";
}

// COMMAND EXECUTION - OS Command Injection
function ping_host() {
    $host = $_GET['host'];
    $output = shell_exec("ping -c 4 " . $host);
    echo "<pre>$output</pre>";
}

// PASSWORD RESET - Weak token
function reset_password($db) {
    $email = $_POST['email'];
    $token = md5(time());
    $db->query("UPDATE users SET reset_token = '$token' WHERE email = '$email'");
    mail($email, "Password Reset", "Reset link: https://app.com/reset?token=$token");
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['action'])) {
    if ($_POST['action'] === 'login') login($db);
    if ($_POST['action'] === 'upload') upload_avatar();
    if ($_POST['action'] === 'reset') reset_password($db);
}

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if (isset($_GET['id'])) show_profile($db);
    if (isset($_GET['delete_id'])) delete_user($db);
    if (isset($_GET['host'])) ping_host();
}
?>
