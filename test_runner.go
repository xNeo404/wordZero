package main

import (
	"fmt"
	"os"
	"os/exec"
)

func main() {
	fmt.Println("Running tests for WordZero project...")

	// 运行pkg包的测试
	fmt.Println("\n=== Testing pkg packages ===")
	cmd := exec.Command("go", "test", "./pkg/...", "-v")
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	err := cmd.Run()
	if err != nil {
		fmt.Printf("pkg tests failed: %v\n", err)
	} else {
		fmt.Println("pkg tests passed")
	}

	// 运行test包的测试
	fmt.Println("\n=== Testing test package ===")
	cmd = exec.Command("go", "test", "./test", "-v")
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	err = cmd.Run()
	if err != nil {
		fmt.Printf("test package failed: %v\n", err)
	} else {
		fmt.Println("test package passed")
	}

	// 运行所有测试
	fmt.Println("\n=== Running all tests ===")
	cmd = exec.Command("go", "test", "./...", "-cover")
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	err = cmd.Run()
	if err != nil {
		fmt.Printf("Overall tests failed: %v\n", err)
		os.Exit(1)
	} else {
		fmt.Println("All tests passed!")
	}
}
