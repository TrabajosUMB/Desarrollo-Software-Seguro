// Package corregido es la versión de ../vulnerable con cada error resuelto.
// Cada función conserva su nombre para poder compararlas lado a lado.
package corregido

import (
	"crypto/rand"
	"crypto/tls"
	"database/sql"
	"encoding/hex"
	"fmt"
	"net/http"
	"os"
	"os/exec"
	"time"

	"golang.org/x/crypto/bcrypt"
)

// dbPassword lee la credencial del entorno, nunca del código fuente.
func dbPassword() string {
	return os.Getenv("DB_PASSWORD")
}

// buscarUsuario usa una consulta parametrizada: el dato viaja aparte de la instrucción.
func buscarUsuario(db *sql.DB, w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	rows, err := db.Query("SELECT name FROM users WHERE id = ?", id)
	if err != nil {
		http.Error(w, "error al consultar", http.StatusInternalServerError)
		return
	}
	defer rows.Close()
	fmt.Fprintln(w, "listo")
}

// hashClave usa bcrypt: lento a propósito y con sal incluida.
func hashClave(clave string) (string, error) {
	h, err := bcrypt.GenerateFromPassword([]byte(clave), bcrypt.DefaultCost)
	return string(h), err
}

// hostsPermitidos es la lista cerrada de destinos que ping puede tocar.
var hostsPermitidos = map[string]bool{"intranet.local": true, "gateway.local": true}

// ping valida contra una lista de permitidos, no usa shell (los argumentos van separados)
// y no devuelve la salida del comando: solo responde si el destino contestó.
func ping(w http.ResponseWriter, r *http.Request) {
	host := r.URL.Query().Get("host")
	if !hostsPermitidos[host] {
		http.Error(w, "destino no permitido", http.StatusForbidden)
		return
	}
	// #nosec G204 G702 -- host validado contra una lista cerrada y sin pasar por una shell
	if err := exec.Command("ping", "-c", "1", host).Run(); err != nil {
		http.Error(w, "el destino no respondió", http.StatusBadGateway)
		return
	}
	fmt.Fprintln(w, "destino accesible")
}

// leerArchivo abre el archivo dentro de una raíz: no puede salir de /var/datos aunque venga ../
func leerArchivo(w http.ResponseWriter, r *http.Request) {
	raiz, err := os.OpenRoot("/var/datos")
	if err != nil {
		http.Error(w, "error", http.StatusInternalServerError)
		return
	}
	defer raiz.Close()
	f, err := raiz.Open(r.URL.Query().Get("f"))
	if err != nil {
		http.Error(w, "no encontrado", http.StatusNotFound)
		return
	}
	defer f.Close()
	http.ServeContent(w, r, "", time.Time{}, f)
}

// clienteSeguro conserva la verificación del certificado y exige TLS 1.2 como mínimo.
func clienteSeguro() *http.Client {
	return &http.Client{Transport: &http.Transport{
		TLSClientConfig: &tls.Config{MinVersion: tls.VersionTLS12},
	}}
}

// tokenSesion usa el generador criptográfico del sistema.
func tokenSesion() (string, error) {
	b := make([]byte, 32)
	if _, err := rand.Read(b); err != nil {
		return "", err
	}
	return hex.EncodeToString(b), nil
}

// iniciarSesion entrega una cookie con Secure, HttpOnly y SameSite.
func iniciarSesion(w http.ResponseWriter, r *http.Request) {
	token, err := tokenSesion()
	if err != nil {
		http.Error(w, "error", http.StatusInternalServerError)
		return
	}
	http.SetCookie(w, &http.Cookie{
		Name:     "sid",
		Value:    token,
		Secure:   true,
		HttpOnly: true,
		SameSite: http.SameSiteStrictMode,
	})
}

// servir escucha solo en la interfaz local y con tiempos límite.
func servir(db *sql.DB) error {
	mux := http.NewServeMux()
	mux.HandleFunc("/buscar", func(w http.ResponseWriter, r *http.Request) { buscarUsuario(db, w, r) })
	mux.HandleFunc("/ping", ping)
	mux.HandleFunc("/archivo", leerArchivo)
	mux.HandleFunc("/login", iniciarSesion)
	srv := &http.Server{
		Addr:              "127.0.0.1:8080",
		Handler:           mux,
		ReadHeaderTimeout: 5 * time.Second,
		ReadTimeout:       10 * time.Second,
		WriteTimeout:      10 * time.Second,
		IdleTimeout:       60 * time.Second,
	}
	return srv.ListenAndServe()
}
