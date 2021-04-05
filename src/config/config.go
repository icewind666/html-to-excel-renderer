package config

import (
	"os"
	"strconv"
	"strings"
)

type SizeTransformConfig struct {
	PxToExcelWidthMultiplier  float64
	PxToExcelHeightMultiplier float64
}

type Config struct {
	SizeTransform SizeTransformConfig
	BatchSize int
	DebugMode bool
	LogLevel string
}

// New returns a new Config struct
func New() *Config {
	return &Config{
		SizeTransform: SizeTransformConfig{
			PxToExcelWidthMultiplier: getEnvAsFloat64("PxToExcelWidthMultiplier", 0.15),
			PxToExcelHeightMultiplier:   getEnvAsFloat64("PxToExcelHeightMultiplier", 0.10),
		},
		DebugMode: getEnvAsBool("DebugMode", true),
		BatchSize: getEnvAsInt("BatchSize", 10000),
		LogLevel: getEnv("GoRenderLogLevel", "info"),
	}
}

// Simple helper function to read an environment or return a default value
func getEnv(key string, defaultVal string) string {
	if value, exists := os.LookupEnv(key); exists {
		return value
	}

	return defaultVal
}

// Simple helper function to read an environment variable into integer or return a default value
func getEnvAsInt(name string, defaultVal int) int {
	valueStr := getEnv(name, "")
	if value, err := strconv.Atoi(valueStr); err == nil {
		return value
	}

	return defaultVal
}

// Simple helper function to read an environment variable into integer or return a default value
func getEnvAsFloat64(name string, defaultVal float64) float64 {
	valueStr := getEnv(name, "")
	if value, err := strconv.ParseFloat(valueStr, 64); err == nil {
		return value
	}

	return defaultVal
}

// Helper to read an environment variable into a bool or return default value
func getEnvAsBool(name string, defaultVal bool) bool {
	valStr := getEnv(name, "")
	if val, err := strconv.ParseBool(valStr); err == nil {
		return val
	}

	return defaultVal
}

// Helper to read an environment variable into a string slice or return default value
func getEnvAsSlice(name string, defaultVal []string, sep string) []string {
	valStr := getEnv(name, "")

	if valStr == "" {
		return defaultVal
	}

	val := strings.Split(valStr, sep)

	return val
}