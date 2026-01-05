//go:build windows

package sugar_test

import (
	"fmt"
	"log"
	"testing"

	"github.com/xll-gen/sugar"
)

// This example demonstrates the standard usage pattern using sugar.Do.

// It automatically handles COM initialization, thread locking, and resource cleanup.

func ExampleDo() {

	sugar.Do(func(ctx *sugar.Context) error {

		excel := ctx.Create("Excel.Application")

		if err := excel.Err(); err != nil {

			// Handle error (e.g., Excel not installed)

			log.Println("Failed to create Excel:", err)

			return err

		}

		// Ensure Excel quits when finished

		defer excel.Call("Quit")



		// Chain methods

		excel.Put("Visible", true)

		

		wb := excel.Get("Workbooks").Call("Add")

		

		name, _ := wb.Get("Name").Value()

		fmt.Printf("Workbook Name: %v\n", name)

		

		// All resources (excel, wb, intermediate objects) are released automatically

		// when the function returns.

		return nil

	})

}



// This example demonstrates getting an active object.

func ExampleGetActive() {

	sugar.Do(func(ctx *sugar.Context) error {

		// This example will fail if Excel is not running, which is expected in a

		// test environment where we don't start it first.

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

	// These tests use a nil dispatch to test internal error handling and logic

	// without requiring a real COM object.

	

	t.Run("Nil Dispatch Error", func(t *testing.T) {

		c := sugar.From(nil)

		err := c.Get("Prop").Err()

		// Calling Get on nil dispatch should return an error.

		if err == nil {

			t.Error("expected error for nil dispatch, got nil")

		}

	})



	t.Run("Error Propagation", func(t *testing.T) {

		c := sugar.From(nil)

		// Manually set an error via a failed operation or just check initial state

		if err := c.Err(); err != nil {

			t.Errorf("initial error should be nil, got %v", err)

		}

	})

	

	t.Run("Create invalid ProgID", func(t *testing.T) {

		// Create requires COM initialization

		sugar.Do(func(ctx *sugar.Context) error {

			c := ctx.Create("Invalid.ProgID.That.Does.Not.Exist")

			if err := c.Err(); err == nil {

				t.Error("expected error for invalid ProgID, got nil")

			}

			return nil

		})

	})

}
