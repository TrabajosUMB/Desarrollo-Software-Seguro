// Package vulnerable reúne errores de seguridad típicos en Go con fines didácticos.
// No está pensado para ejecutarse: solo para ser analizado con gosec.
package vulnerable

import (
	"crypto/md5"
	"crypto/tls"
	"database/sql"
	"encoding/hex"
	"fmt"
	"math/rand"
	"net/http"
	"os"
	"os/exec"
)

// Credencial escrita en el código fuente (G101).
const dbPassword = "x9Qm2LvT8rZkP4wYcN7bHs3JdFa6GuE1"

// buscarUsuario arma la consulta SQL concatenando el parámetro de la petición (G201, G701).
func buscarUsuario(db *sql.DB, w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	consulta := fmt.Sprintf("SELECT name FROM users WHERE id = %s", id)
	rows, _ := db.Query(consulta)
	defer rows.Close()
	fmt.Fprintln(w, "listo")
}

// hashClave protege contraseñas con MD5, un algoritmo roto (G401, G501).
func hashClave(clave string) string {
	h := md5.Sum([]byte(clave))
	return hex.EncodeToString(h[:])
}

// ping pasa texto del usuario a una shell (G204, G702).
func ping(w http.ResponseWriter, r *http.Request) {
	host := r.URL.Query().Get("host")
	salida, _ := exec.Command("sh", "-c", "ping -c 1 "+host).Output()
	w.Write(salida)
}

// leerArchivo une una ruta fija con un nombre que escribe el usuario (G304, G703).
func leerArchivo(w http.ResponseWriter, r *http.Request) {
	nombre := r.URL.Query().Get("f")
	datos, _ := os.ReadFile("/var/datos/" + nombre)
	w.Write(datos)
}

// clienteInseguro desactiva la verificación del certificado TLS (G402).
func clienteInseguro() *http.Client {
	return &http.Client{Transport: &http.Transport{
		TLSClientConfig: &tls.Config{InsecureSkipVerify: true},
	}}
}

// tokenSesion genera un identificador de sesión con un generador no criptográfico (G404).
func tokenSesion() string {
	return fmt.Sprintf("%d", rand.Int63())
}

// iniciarSesion entrega una cookie sin Secure, HttpOnly ni SameSite (G124).
func iniciarSesion(w http.ResponseWriter, r *http.Request) {
	http.SetCookie(w, &http.Cookie{Name: "sid", Value: tokenSesion()})
}

// servir escucha en todas las interfaces y sin tiempos límite (G102, G114).
func servir(db *sql.DB) error {
	http.HandleFunc("/buscar", func(w http.ResponseWriter, r *http.Request) { buscarUsuario(db, w, r) })
	http.HandleFunc("/ping", ping)
	http.HandleFunc("/archivo", leerArchivo)
	http.HandleFunc("/login", iniciarSesion)
	return http.ListenAndServe(":8080", nil)
}
