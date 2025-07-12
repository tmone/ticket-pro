import '@testing-library/jest-dom'

// Mock navigator.mediaDevices for camera tests
Object.defineProperty(navigator, 'mediaDevices', {
  writable: true,
  value: {
    getUserMedia: jest.fn(),
  },
})

// Mock URL.createObjectURL
global.URL.createObjectURL = jest.fn(() => 'mock-url')
global.URL.revokeObjectURL = jest.fn()

// Mock FileReader
global.FileReader = class {
  constructor() {
    this.result = null
    this.onload = null
    this.onerror = null
  }
  
  readAsArrayBuffer(file) {
    // Simulate async file reading
    setTimeout(() => {
      if (this.onload) {
        this.result = new ArrayBuffer(8)
        this.onload({ target: { result: this.result } })
      }
    }, 0)
  }
}

// Mock requestAnimationFrame
global.requestAnimationFrame = jest.fn(cb => setTimeout(cb, 16))
global.cancelAnimationFrame = jest.fn()