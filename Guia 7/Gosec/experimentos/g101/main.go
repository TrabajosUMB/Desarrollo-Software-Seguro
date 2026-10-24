package main

import "fmt"

const claveLegible = "Sup3rS3cret!2026"                   // contraseña "humana", baja entropía
const claveAleatoria = "x9Qm2LvT8rZkP4wYcN7bHs3JdFa6GuE1" // 32 caracteres aleatorios
var password = "admin123"                                 // el clásico
var apiToken = "K7f9Zp2QmX4vB8nL1cR6tY3wH5jD0sGa"

func main() {
	fmt.Println(len(claveLegible), len(claveAleatoria), len(password), len(apiToken))
}
