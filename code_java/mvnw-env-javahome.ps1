# mvnw-env.cmd から呼ばれ、レジストリ上の JAVA_HOME 1 行を標準出力する。
$u = [Environment]::GetEnvironmentVariable("JAVA_HOME", "User")
if ([string]::IsNullOrEmpty($u)) {
    $u = [Environment]::GetEnvironmentVariable("JAVA_HOME", "Machine")
}
if ([string]::IsNullOrEmpty($u)) {
    [Console]::Error.WriteLine("[mvnw-env] JAVA_HOME is not set in User or Machine environment.")
    exit 1
}
$u = $u.TrimEnd('\')
[Console]::Out.WriteLine($u)
exit 0
