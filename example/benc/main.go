package main

import (
	"fmt"
	"runtime"
	"time"

	"github.com/didikprabowo/mbadocx"
)

type MemWatcher struct {
	interval time.Duration
	stop     chan struct{}
	peakMB   float64
}

func NewMemWatcher(interval time.Duration) *MemWatcher {
	return &MemWatcher{
		interval: interval,
		stop:     make(chan struct{}),
	}
}

func (mw *MemWatcher) Start() {
	go func() {
		ticker := time.NewTicker(mw.interval)
		defer ticker.Stop()
		for {
			select {
			case <-ticker.C:
				var m runtime.MemStats
				runtime.ReadMemStats(&m)
				current := float64(m.Alloc) / (1024 * 1024)
				if current > mw.peakMB {
					mw.peakMB = current
				}
				fmt.Printf("MEM: Alloc = %.2f MiB | Sys = %.2f MiB | GC = %d\n",
					current,
					float64(m.Sys)/(1024*1024),
					m.NumGC)
			case <-mw.stop:
				return
			}
		}
	}()
}

func (mw *MemWatcher) Stop() {
	close(mw.stop)
}

func (mw *MemWatcher) Peak() float64 {
	return mw.peakMB
}

func main() {
	write()
}

func write() {
	watcher := NewMemWatcher(200 * time.Millisecond)
	watcher.Start()

	doc := mbadocx.New()
	defer doc.Close()

	for i := 0; i < 20000; i++ {
		doc.AddParagraph().AddText("Hello world")
		if i%5000 == 0 {
			time.Sleep(1 * time.Second) // simulate heavy work
		}
	}
	// heavy work ...
	_ = doc.Save("testdata/big.docx")

	watcher.Stop()
	fmt.Printf("Peak memory usage: %.2f MiB\n", watcher.Peak())
}
