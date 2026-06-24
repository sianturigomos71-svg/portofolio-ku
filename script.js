// Page Loader
window.addEventListener('load', () => {
    const loader = document.getElementById('page-loader');
    const percentage = document.querySelector('.loader-percentage');
    const progress = document.querySelector('.loader-progress');
    let progressValue = 0;

    const interval = setInterval(() => {
        progressValue += Math.floor(Math.random() * 10) + 1;
        if (progressValue >= 100) {
            progressValue = 100;
            clearInterval(interval);
            setTimeout(() => {
                loader.classList.add('hidden');
                // Trigger hero text animations after loader
                initHeroTextAnimations();
            }, 500);
        }
        percentage.textContent = progressValue + '%';
        progress.style.width = progressValue + '%';
    }, 100);
});

// Smooth Scrolling
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    });
});

// Lightbox Functionality
const lightbox = document.getElementById('lightbox');
const lightboxImg = document.getElementById('lightbox-img');
const lightboxClose = document.querySelector('.lightbox-close');

// Add click event to all images
document.querySelectorAll('img').forEach(img => {
    img.addEventListener('click', function() {
        lightbox.style.display = 'flex';
        lightboxImg.src = this.src;
        document.body.style.overflow = 'hidden';
    });
});

// Close lightbox when clicking close button
lightboxClose.addEventListener('click', function() {
    lightbox.style.display = 'none';
    document.body.style.overflow = 'auto';
});

// Close lightbox when clicking outside the image
lightbox.addEventListener('click', function(e) {
    if (e.target === lightbox) {
        lightbox.style.display = 'none';
        document.body.style.overflow = 'auto';
    }
});

// Close lightbox with Escape key
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && lightbox.style.display === 'flex') {
        lightbox.style.display = 'none';
        document.body.style.overflow = 'auto';
    }
});

// PDF Modal Functionality
const pdfModal = document.getElementById('pdfModal');
const pdfFrame = document.getElementById('pdfFrame');
const pdfModalClose = document.querySelector('.pdf-modal-close');

function openPdfModal(pdfPath) {
    pdfFrame.src = pdfPath;
    pdfModal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

// Close PDF modal when clicking close button
pdfModalClose.addEventListener('click', function() {
    pdfModal.style.display = 'none';
    pdfFrame.src = '';
    document.body.style.overflow = 'auto';
});

// Close PDF modal when clicking outside the content
pdfModal.addEventListener('click', function(e) {
    if (e.target === pdfModal) {
        pdfModal.style.display = 'none';
        pdfFrame.src = '';
        document.body.style.overflow = 'auto';
    }
});

// Close PDF modal with Escape key
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && pdfModal.style.display === 'block') {
        pdfModal.style.display = 'none';
        pdfFrame.src = '';
        document.body.style.overflow = 'auto';
    }
});

// Gallery Modal Functionality
const galleryModal = document.getElementById('galleryModal');
const galleryImage = document.getElementById('galleryImage');
const galleryModalClose = document.querySelector('.gallery-modal-close');
const galleryPrev = document.querySelector('.gallery-prev');
const galleryNext = document.querySelector('.gallery-next');
const galleryCounter = document.querySelector('.gallery-counter');

let currentGalleryImages = [];
let currentImageIndex = 0;

function openGalleryModal(images) {
    currentGalleryImages = images;
    currentImageIndex = 0;
    galleryImage.src = images[0];
    updateGalleryCounter();
    galleryModal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

function updateGalleryCounter() {
    galleryCounter.textContent = `${currentImageIndex + 1} / ${currentGalleryImages.length}`;
}

function showNextImage() {
    currentImageIndex = (currentImageIndex + 1) % currentGalleryImages.length;
    galleryImage.src = currentGalleryImages[currentImageIndex];
    updateGalleryCounter();
}

function showPrevImage() {
    currentImageIndex = (currentImageIndex - 1 + currentGalleryImages.length) % currentGalleryImages.length;
    galleryImage.src = currentGalleryImages[currentImageIndex];
    updateGalleryCounter();
}

// Close gallery modal when clicking close button
galleryModalClose.addEventListener('click', function() {
    galleryModal.style.display = 'none';
    galleryImage.src = '';
    document.body.style.overflow = 'auto';
});

// Navigate to next image
galleryNext.addEventListener('click', function(e) {
    e.stopPropagation();
    showNextImage();
});

// Navigate to previous image
galleryPrev.addEventListener('click', function(e) {
    e.stopPropagation();
    showPrevImage();
});

// Close gallery modal when clicking outside the content
galleryModal.addEventListener('click', function(e) {
    if (e.target === galleryModal) {
        galleryModal.style.display = 'none';
        galleryImage.src = '';
        document.body.style.overflow = 'auto';
    }
});

// Close gallery modal with Escape key
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && galleryModal.style.display === 'block') {
        galleryModal.style.display = 'none';
        galleryImage.src = '';
        document.body.style.overflow = 'auto';
    }
    // Navigate with arrow keys
    if (galleryModal.style.display === 'block') {
        if (e.key === 'ArrowRight') {
            showNextImage();
        } else if (e.key === 'ArrowLeft') {
            showPrevImage();
        }
    }
});

// Split Text Animation for Hero
function initHeroTextAnimations() {
    const h1 = document.querySelector('.hero h1');
    const h2 = document.querySelector('.hero h2');

    if (h1) {
        const text = h1.textContent;
        h1.innerHTML = '';
        text.split('').forEach((char, index) => {
            const span = document.createElement('span');
            span.className = 'char';
            span.textContent = char === ' ' ? '\u00A0' : char;
            span.style.animationDelay = `${index * 0.05}s`;
            h1.appendChild(span);
        });
    }

    if (h2) {
        const text = h2.textContent;
        h2.innerHTML = '';
        text.split(' ').forEach((word, index) => {
            const span = document.createElement('span');
            span.className = 'word';
            span.textContent = word;
            span.style.animationDelay = `${index * 0.15}s`;
            h2.appendChild(span);
            if (index < text.split(' ').length - 1) {
                h2.appendChild(document.createTextNode(' '));
            }
        });
    }
}

// Sidebar Toggle
const hamburgerBtn = document.getElementById('hamburgerBtn');
const sidebar = document.getElementById('sidebar');
const closeBtn = document.getElementById('closeBtn');

hamburgerBtn.addEventListener('click', () => {
    sidebar.classList.toggle('active');
});

closeBtn.addEventListener('click', () => {
    sidebar.classList.remove('active');
});

// Close sidebar when clicking on a menu item
document.querySelectorAll('.sidebar-menu a').forEach(link => {
    link.addEventListener('click', () => {
        sidebar.classList.remove('active');
    });
});

// Close sidebar when clicking outside
document.addEventListener('click', (e) => {
    if (!sidebar.contains(e.target) && !hamburgerBtn.contains(e.target)) {
        sidebar.classList.remove('active');
    }
});

// Tab Navigation for Artefak and Lampiran Sections
const tabBtns = document.querySelectorAll('.tab-btn');

tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        // Find the parent section of the clicked button
        const parentSection = btn.closest('section');
        
        // Remove active class from all buttons and contents within this section only
        parentSection.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        parentSection.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        
        // Add active class to clicked button
        btn.classList.add('active');
        
        // Show corresponding tab content within this section
        const tabId = btn.getAttribute('data-tab');
        document.getElementById(tabId).classList.add('active');
    });
});

// Scroll Spy Navigation for Sidebar
const sections = document.querySelectorAll('section');
const sidebarLinks = document.querySelectorAll('.sidebar-menu a');

window.addEventListener('scroll', () => {
    let current = '';
    
    sections.forEach(section => {
        const sectionTop = section.offsetTop;
        const sectionHeight = section.clientHeight;
        
        if (pageYOffset >= sectionTop - 200) {
            current = section.getAttribute('id');
        }
    });
    
    sidebarLinks.forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href') === `#${current}`) {
            link.classList.add('active');
        }
    });
});

// Mobile Navigation Toggle
const hamburger = document.querySelector('.hamburger');
const navMenu = document.querySelector('.nav-menu');

hamburger.addEventListener('click', () => {
    hamburger.classList.toggle('active');
    navMenu.classList.toggle('active');
});

// Close mobile menu when clicking on a link
document.querySelectorAll('.nav-menu a').forEach(link => {
    link.addEventListener('click', () => {
        hamburger.classList.remove('active');
        navMenu.classList.remove('active');
    });
});

// Smooth scroll for navigation links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    });
});

// Navbar background change on scroll
const navbar = document.querySelector('.navbar');
window.addEventListener('scroll', () => {
    if (window.scrollY > 50) {
        navbar.style.boxShadow = '0 4px 6px -1px rgba(0, 0, 0, 0.1)';
    } else {
        navbar.style.boxShadow = 'none';
    }
});

// Contact Form Submission
const contactForm = document.getElementById('contactForm');
contactForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    // Get form data
    const formData = new FormData(contactForm);
    
    // Show success message (in a real application, you would send this to a server)
    alert('Terima kasih! Pesan Anda telah terkirim. Kami akan menghubungi Anda segera.');
    
    // Reset form
    contactForm.reset();
});

// Animate skill bars on scroll
const skillBars = document.querySelectorAll('.skill-progress');
const animateSkillBars = () => {
    skillBars.forEach(bar => {
        const barTop = bar.getBoundingClientRect().top;
        const windowHeight = window.innerHeight;
        
        if (barTop < windowHeight * 0.8) {
            bar.style.width = bar.style.width;
        }
    });
};

window.addEventListener('scroll', animateSkillBars);
window.addEventListener('load', animateSkillBars);

// Add animation on scroll for cards
const observerOptions = {
    threshold: 0.1,
    rootMargin: '0px 0px -50px 0px'
};

const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
        if (entry.isIntersecting) {
            entry.target.style.opacity = '1';
            entry.target.style.transform = 'translateY(0)';
        }
    });
}, observerOptions);

// Observe all cards
document.querySelectorAll('.experience-card, .achievement-card, .document-card').forEach(card => {
    card.style.opacity = '0';
    card.style.transform = 'translateY(20px)';
    card.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
    observer.observe(card);
});

// Typing effect for hero section (optional enhancement)
const heroTitle = document.querySelector('.hero-text h1');
if (heroTitle) {
    const originalText = heroTitle.textContent;
    heroTitle.textContent = '';
    let i = 0;
    
    const typeWriter = () => {
        if (i < originalText.length) {
            heroTitle.textContent += originalText.charAt(i);
            i++;
            setTimeout(typeWriter, 100);
        }
    };
    
    // Start typing effect after a short delay
    setTimeout(typeWriter, 500);
}

// Active navigation link highlighting
const navLinks = document.querySelectorAll('.nav-menu a');

window.addEventListener('scroll', () => {
    let current = '';
    
    sections.forEach(section => {
        const sectionTop = section.offsetTop;
        const sectionHeight = section.clientHeight;
        
        if (pageYOffset >= sectionTop - 200) {
            current = section.getAttribute('id');
        }
    });
    
    navLinks.forEach(link => {
        link.style.color = '';
        if (link.getAttribute('href') === `#${current}`) {
            link.style.color = 'var(--primary-color)';
        }
    });
});

// Document card button interactions
document.querySelectorAll('.document-card .btn').forEach(button => {
    button.addEventListener('click', (e) => {
        e.preventDefault();
        const documentName = button.parentElement.querySelector('h3').textContent;
        alert(`Dokumen "${documentName}" akan dibuka. (Dalam implementasi nyata, ini akan membuka file PDF atau dokumen lainnya)`);
    });
});

// Add hover effect to social links
document.querySelectorAll('.social-links a').forEach(link => {
    link.addEventListener('mouseenter', function() {
        this.style.transform = 'translateY(-5px) rotate(360deg)';
    });
    
    link.addEventListener('mouseleave', function() {
        this.style.transform = 'translateY(0) rotate(0deg)';
    });
});

// Console message for developers
console.log('E-Portofolio PPG Prajabatan loaded successfully!');
console.log('Built with HTML, CSS, and JavaScript');

// Scroll animations for sections
const animateOnScroll = () => {
    const elements = document.querySelectorAll('.fade-in, .slide-in-left, .slide-in-right, .scale-in');
    
    elements.forEach(element => {
        const elementTop = element.getBoundingClientRect().top;
        const windowHeight = window.innerHeight;
        
        if (elementTop < windowHeight * 0.85) {
            element.style.opacity = '1';
            element.style.transform = 'translateY(0) translateX(0) scale(1)';
        }
    });
};

window.addEventListener('scroll', animateOnScroll);
window.addEventListener('load', animateOnScroll);

// Menu items staggered animation
const menuItems = document.querySelectorAll('.menu-item');
menuItems.forEach((item, index) => {
    item.style.opacity = '0';
    item.style.transform = 'translateY(20px)';
    item.style.transition = 'opacity 0.5s ease, transform 0.5s ease';
    
    setTimeout(() => {
        item.style.opacity = '1';
        item.style.transform = 'translateY(0)';
    }, 500 + (index * 100));
});

// Menu item click animations and active state
menuItems.forEach(item => {
    item.addEventListener('click', function(e) {
        // Remove active class from all menu items
        menuItems.forEach(mi => mi.classList.remove('active'));
        
        // Add active class to clicked item
        this.classList.add('active');
        
        // Add click animation
        this.style.transform = 'scale(0.9)';
        setTimeout(() => {
            this.style.transform = '';
        }, 150);
    });
});

// Active menu item based on scroll position
window.addEventListener('scroll', () => {
    let current = '';
    
    sections.forEach(section => {
        const sectionTop = section.offsetTop;
        const sectionHeight = section.clientHeight;
        
        if (pageYOffset >= sectionTop - 200) {
            current = section.getAttribute('id');
        }
    });
    
    menuItems.forEach(item => {
        item.classList.remove('active');
        const href = item.getAttribute('href');
        if (href === `#${current}`) {
            item.classList.add('active');
        }
    });
});

// Instant scroll (teleportation effect)
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            // Instant scroll without animation
            target.scrollIntoView({
                behavior: 'auto',
                block: 'start'
            });
        }
    });
});

// File Upload Handling for RPP
document.querySelectorAll('.upload-input').forEach(input => {
    input.addEventListener('change', function(e) {
        const file = e.target.files[0];
        const statusElement = this.parentElement.querySelector('.upload-status');
        
        if (file) {
            statusElement.textContent = `File terpilih: ${file.name}`;
            statusElement.style.color = '#00ff88';
        } else {
            statusElement.textContent = 'Belum ada file diupload';
            statusElement.style.color = '#a0a0a0';
        }
    });
});
