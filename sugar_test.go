//go:build windows

package sugar_test

import (
	"fmt"
	"log"
	"testing"

	"github.com/xll-gen/sugar"
)

func ExampleDo() {
	sugar.Do(func(ctx sugar.Context) error {
		excel := ctx.Create("Excel.Application")
		if err := excel.Err(); err != nil {
			log.Println("Failed to create Excel:", err)
			return err
		}
		defer excel.Call("Quit")

		excel.Put("Visible", true)
		wb := excel.Get("Workbooks").Call("Add")
		name, _ := wb.Get("Name").Value()
		fmt.Printf("Workbook Name: %v\n", name)
		return nil
	})
}

func ExampleGetActive() {
	sugar.Do(func(ctx sugar.Context) error {
		excel := ctx.GetActive("Excel.Application")
		if err := excel.Err(); err != nil {
			fmt.Println("GetActive failed as expected.")
		} else {
			fmt.Println("GetActive succeeded.")
		}
		return nil
	})
}

func TestChain_Mock(t *testing.T) {
	t.Run("Nil Dispatch Error", func(t *testing.T) {
		c := sugar.From(nil)
		err := c.Get("Prop").Err()
		if err == nil {
			t.Error("expected error for nil dispatch, got nil")
		}
	})

	t.Run("Error Propagation", func(t *testing.T) {
		c := sugar.From(nil)
		if err := c.Err(); err != nil {
			t.Errorf("initial error should be nil, got %v", err)
		}
	})
	
	t.Run("Create invalid ProgID", func(t *testing.T) {
		sugar.Do(func(ctx sugar.Context) error {
			c := ctx.Create("Invalid.ProgID.That.Does.Not.Exist")
			if err := c.Err(); err == nil {
				t.Error("expected error for invalid ProgID, got nil")
			}
			return nil
		})
	})
}