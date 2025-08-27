package utils

import (
	"io"
	"os"
)

// CopyFile 复制文件，保持所有格式
func CopyFile(srcPath, dstPath string) error {
	// 使用文件系统复制，保持所有格式
	srcFile, err := os.Open(srcPath)
	if err != nil {
		return err
	}
	defer srcFile.Close()

	dstFile, err := os.Create(dstPath)
	if err != nil {
		return err
	}
	defer dstFile.Close()

	// 复制文件内容
	_, err = io.Copy(dstFile, srcFile)
	return err
}
