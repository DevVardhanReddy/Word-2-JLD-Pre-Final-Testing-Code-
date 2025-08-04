const menuToggle = document.getElementById('menuToggle');
const navContent = document.getElementById('navContent');

menuToggle.addEventListener('click', () => {
  menuToggle.classList.toggle('open');
  navContent.classList.toggle('active');
});
  
  const cards = document.querySelectorAll('.card');

  cards.forEach(card => {
    card.addEventListener('mousemove', e => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;

      card.style.setProperty('--x', `${x}px`);
      card.style.setProperty('--y', `${y}px`);
    });

    card.addEventListener('mouseleave', () => {
      // Optional: reset to center or keep last
      card.style.setProperty('--x', `50%`);
      card.style.setProperty('--y', `50%`);
    });
  });

  
document.querySelectorAll('.scroll-card').forEach(card => {
    gsap.to(card, {
      scale: 0.7,
      opacity: 0,
      scrollTrigger: {
        trigger: card,
        start: "top 0%",
        end: "bottom 20%",
        scrub: true,
        // markers: true
      }
    });
  });




const words = [
  'Enterprise Growth',
  'Digital Transformation',
  'Business Intelligence',
  'Document Automation',
  'Workflow Efficiency',
  'Data Analytics',
  'Process Innovation'
];

let currentWordIndex = 0;
let currentCharIndex = 0;
let isDeleting = false;
const typewriterElement = document.getElementById('typewriter');
const typeSpeed = 80;           // reduced from 100
const deleteSpeed = 100;         // reduced from 50
const pauseAfterWord = 1300;    // reduced from 1500
const pauseAfterDelete = 200;   // reduced from 300

function typeWriter() {
  const currentWord = words[currentWordIndex];
  
  if (!isDeleting) {
      if (currentCharIndex < currentWord.length) {
          typewriterElement.textContent = currentWord.substring(0, currentCharIndex + 1);
          currentCharIndex++;
          setTimeout(typeWriter, typeSpeed);
      } else {
          setTimeout(() => {
              isDeleting = true;
              typeWriter();
          }, pauseAfterWord);
      }
  } else {
      if (currentCharIndex > 0) {
          typewriterElement.textContent = currentWord.substring(0, currentCharIndex - 1);
          currentCharIndex--;
          setTimeout(typeWriter, deleteSpeed);
      } else {
          isDeleting = false;
          currentWordIndex = (currentWordIndex + 1) % words.length;
          setTimeout(typeWriter, pauseAfterDelete);
      }
  }
}

setTimeout(typeWriter, 1000);


function updateHighlights() {
const container = document.querySelector('.scrolling-container');
const items = container.querySelectorAll('.item');
const midPoint = container.getBoundingClientRect().top + container.offsetHeight / 2;

let distances = [];

items.forEach(item => {
  const itemCenter = item.getBoundingClientRect().top + item.offsetHeight / 2;
  const distance = Math.abs(midPoint - itemCenter);
  distances.push({ item, distance });
});

distances.sort((a, b) => a.distance - b.distance);

// Reset all styles
items.forEach(i => {
  i.classList.remove('highlight', 'mid-dim');
  i.style.color = 'rgba(255,255,255,0.3)';
});

// Apply highlight and mid-dim
if (distances[0]) distances[0].item.classList.add('highlight');
if (distances[1]) distances[1].item.classList.add('highlight');
if (distances[2]) distances[2].item.classList.add('mid-dim');
if (distances[3]) distances[3].item.classList.add('mid-dim');
}

setInterval(updateHighlights, 100);
window.addEventListener('load', updateHighlights);